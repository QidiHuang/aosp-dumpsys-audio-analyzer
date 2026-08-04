import re

with open('dumpsys_files/music_play.txt', 'r') as f:
    content = f.read()

# Threads
threads = list(re.finditer(r'Output thread 0x[0-9a-fA-F]+', content))
print("Threads found:", len(threads))
if threads:
    print(threads[0].group(0))

# Standby
standby = list(re.finditer(r'Standby:\s*(yes|no)', content))
print("Standby found:", len(standby))

# Sink buffer
sink = list(re.finditer(r'Sink buffer\s*:\s*\d+\s*frames', content))
print("Sink buffer found:", len(sink))

# Track Info
# Usually looks like:
#   Id Active Client Session Port Id S  Flags   Format Chn mask  SRate mFmMk     Server fCount     Active fCount
#   0    yes   2333       0       0   A  0x000 0x00000001 0x0003  48000 00000000 00000000  00000000 00000000 00000000

track_clients = list(re.finditer(r'^\s*(\w+)\s+(yes|no)\s+(\d+)\s+', content, re.MULTILINE))
print("Tracks found:", len(track_clients))
if track_clients:
    print(track_clients[0].groups())
