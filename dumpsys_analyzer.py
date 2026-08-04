#!/usr/bin/env python3
import tkinter as tk
from tkinter import ttk
import subprocess
import re
import sys
import argparse
import time

class DumpsysAnalyzerApp:
    def __init__(self, root, file_source=None):
        self.root = root
        self.root.title("Dumpsys Analyzer")
        self.root.geometry("1000x700")

        self.file_source = file_source
        self.running = True

        # Top panel for controls
        self.control_frame = tk.Frame(self.root)
        self.control_frame.pack(fill=tk.X, side=tk.BOTTOM, pady=5, padx=5)

        # Main Canvas with scrollbars
        self.canvas_frame = tk.Frame(self.root)
        self.canvas_frame.pack(fill=tk.BOTH, expand=True)

        self.canvas = tk.Canvas(self.canvas_frame, bg="white")

        self.v_scroll = ttk.Scrollbar(self.canvas_frame, orient=tk.VERTICAL, command=self.canvas.yview)
        self.v_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        self.h_scroll = ttk.Scrollbar(self.canvas_frame, orient=tk.HORIZONTAL, command=self.canvas.xview)
        self.h_scroll.pack(side=tk.BOTTOM, fill=tk.X)

        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.canvas.config(yscrollcommand=self.v_scroll.set, xscrollcommand=self.h_scroll.set)

        # Mouse wheel scrolling
        self.canvas.bind("<MouseWheel>", self._on_mousewheel)
        self.canvas.bind("<Button-4>", self._on_mousewheel)
        self.canvas.bind("<Button-5>", self._on_mousewheel)

    def _on_mousewheel(self, event):
        if event.num == 4 or event.delta > 0:
            self.canvas.yview_scroll(-1, "units")
        elif event.num == 5 or event.delta < 0:
            self.canvas.yview_scroll(1, "units")

        self.refresh_btn = tk.Button(self.control_frame, text="Refresh", command=self.do_refresh)
        self.refresh_btn.pack(side=tk.LEFT, padx=5)

        self.status_var = tk.StringVar()
        self.status_var.set("Ready")
        self.status_label = tk.Label(self.control_frame, textvariable=self.status_var)
        self.status_label.pack(side=tk.LEFT, padx=10)

        # First draw
        self.do_refresh()

    def do_refresh(self):
        self.status_var.set("Fetching data...")
        self.root.update()

        try:
            data = self.gather_data()
            self.status_var.set("Parsing data...")
            self.root.update()

            parsed = self.parse_data(data)
            self.status_var.set("Drawing UI...")
            self.root.update()

            self.draw_ui(parsed)
            self.status_var.set("Ready")
        except Exception as e:
            self.status_var.set(f"Error: {str(e)}")
            self.draw_ui(None)

    def gather_data(self):
        """
        Gathers data from offline file or directly from ADB.
        Combines dumpsys and ps in one shell execution for efficiency.
        """
        if self.file_source:
            with open(self.file_source, 'r', encoding='utf8', errors='ignore') as f:
                return f.read()
        else:
            # Combine commands: dumpsys media.audio_flinger + ps output
            cmd = "adb shell \"dumpsys media.audio_flinger && echo '---PS_START---' && (ps -A || ps)\""
            try:
                result = subprocess.run(cmd, shell=True, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=10)
                return result.stdout.decode('utf8', errors='ignore')
            except subprocess.TimeoutExpired:
                raise Exception("ADB command timed out. Is device connected?")
            except subprocess.CalledProcessError as e:
                raise Exception(f"ADB command failed: {e.stderr.decode('utf8', errors='ignore')}")

    def parse_data(self, raw_data):
        parts = raw_data.split('---PS_START---')
        dumpsys_out = parts[0]
        ps_out = parts[1] if len(parts) > 1 else ""

        pid_map = {}
        for line in ps_out.splitlines():
            line = line.strip()
            if not line or 'USER' in line or 'PID' in line:
                continue
            tokens = line.split()
            if len(tokens) >= 8:
                pid = tokens[1]
                name = tokens[-1]
                pid_map[pid] = name

        parsed_info = {
            'threads': [],
            'devices': []
        }

        thread_pattern = re.compile(r'(Output|Input) thread 0x([0-9a-fA-F]+)(.*?)(?=(?:Output|Input) thread 0x|$)', re.DOTALL)
        for thread_match in thread_pattern.finditer(dumpsys_out):
            thread_type = thread_match.group(1)
            thread_id = thread_match.group(2)
            block = thread_match.group(3)

            thread_info = {
                'id': thread_id,
                'type': thread_type,
                'standby': 'no',
                'tracks': [],
                'hal_frames': '',
                'frame_size': '',
                'sink_buffer': ''
            }

            m_standby = re.search(r'Standby:\s*(yes|no)', block)
            if m_standby:
                thread_info['standby'] = m_standby.group(1)

            m_hal = re.search(r'HAL frame count:\s*(\d+)', block)
            if m_hal:
                thread_info['hal_frames'] = m_hal.group(1)

            m_size = re.search(r'Processing frame size:\s*(\d+)', block)
            if m_size:
                thread_info['frame_size'] = m_size.group(1)

            m_sink = re.search(r'Sink buffer\s*:\s*(0x[0-9a-fA-F]+)', block)
            if m_sink:
                thread_info['sink_buffer'] = m_sink.group(1)

            # Look for headers like: Type     Id Active Client   OR   Active     Id Client
            track_section_m = re.search(r'(?:Type\s+Id\s+Active\s+Client|Active\s+Id\s+Client).*?(?=\n\s*\n|\n\s*Effect Chains|\n\s*Local log)', block, re.DOTALL)
            if track_section_m:
                track_lines = track_section_m.group(0).splitlines()[1:]
                for line in track_lines:
                    line = line.strip()
                    if not line or 'Effect Chains' in line: break

                    # Capture regular track row: (Type optionally, ID, Active, Client)
                    m_trk = re.match(r'^(?:(\S+)\s+)?(\d+)\s+(yes|no)\s+(\d+)\s+', line)
                    if not m_trk:
                        # Capture input track row: (Active, ID, Client)
                        m_trk = re.match(r'^(yes|no)\s+(\d+)\s+(\d+)\s+', line)
                        if m_trk:
                            t_active = m_trk.group(1)
                            t_id = m_trk.group(2)
                            t_client = m_trk.group(3)
                        else:
                            continue
                    else:
                        t_id = m_trk.group(2)
                        t_active = m_trk.group(3)
                        t_client = m_trk.group(4)

                    app_name = pid_map.get(t_client, f"PID:{t_client}")

                    thread_info['tracks'].append({
                        'id': t_id,
                        'active': t_active,
                        'client_pid': t_client,
                        'app_name': app_name
                    })

            parsed_info['threads'].append(thread_info)

        return parsed_info

    def draw_ui(self, parsed_data):
        self.canvas.delete("all")
        if not parsed_data:
            return

        # Layout columns: App -> Thread -> Track -> Buffer
        col_x = [50, 300, 550, 800]

        apps = {}
        app_nodes = {}

        y_offset_thread = 50

        for th in parsed_data['threads']:
            th_id = th['id']
            th_type = th.get('type', 'Output')

            th_text = f"{th_type} Thread\n0x{th_id}\nStandby: {th['standby']}"
            th_color = "lightgray" if th['standby'] == 'yes' else "lightgreen"

            th_box = self.draw_box(col_x[1], y_offset_thread, th_text, bg=th_color)

            y_offset_track = y_offset_thread

            for tr in th['tracks']:
                pid = tr['client_pid']
                apps[pid] = tr['app_name']

                tr_text = f"Track {tr['id']}\nActive: {tr['active']}"
                tr_color = "lightyellow" if tr['active'] == 'yes' else "white"
                tr_box = self.draw_box(col_x[2], y_offset_track, tr_text, bg=tr_color)

                self.draw_line(th_box, tr_box)

                if pid not in app_nodes:
                    app_nodes[pid] = []
                app_nodes[pid].append(tr_box)

                y_offset_track += 65

            if th['sink_buffer']:
                buf_text = f"Buffer\n{th['sink_buffer']}"
                if th['frame_size']:
                    buf_text += f"\nSize: {th['frame_size']}"
                buf_box = self.draw_box(col_x[3], y_offset_thread, buf_text, bg="cyan")
                self.draw_line(th_box, buf_box)

            y_offset_thread = max(y_offset_thread + 80, y_offset_track + 10)

        y_offset_app = 50
        for pid, name in apps.items():
            if pid == "no" or name == "PID:no":
                app_text = "No App"
            else:
                app_text = f"App\n{name}\n(PID: {pid})"
            app_box = self.draw_box(col_x[0], y_offset_app, app_text, bg="lightblue")

            for tr_box in app_nodes.get(pid, []):
                self.draw_line(app_box, tr_box)

            y_offset_app += 80

        # Update scrolling region
        bbox = self.canvas.bbox("all")
        if bbox:
            self.canvas.config(scrollregion=(0, 0, bbox[2] + 100, bbox[3] + 100))

    def draw_box(self, x, y, text, bg="white", width=150, height=50):
        self.canvas.create_rectangle(x, y, x+width, y+height, fill=bg, outline="black", width=2)
        self.canvas.create_text(x + width/2, y + height/2, text=text, justify=tk.CENTER)
        return (x+width, y+height/2, x, y+height/2) # (out_x, out_y, in_x, in_y)

    def draw_line(self, node_from, node_to):
        fx, fy = node_from[0], node_from[1]
        tx, ty = node_to[2], node_to[3]

        # Draw a bezier-like curve or angled line
        mid_x = (fx + tx) / 2
        self.canvas.create_line(fx, fy, mid_x, fy, mid_x, ty, tx, ty, arrow=tk.LAST, width=2, fill="gray")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Android AudioFlinger Dumpsys Analyzer")
    parser.add_argument("--test-file", help="Read from offline file instead of ADB", type=str)
    args = parser.parse_args()

    root = tk.Tk()
    app = DumpsysAnalyzerApp(root, file_source=args.test_file)
    root.mainloop()
