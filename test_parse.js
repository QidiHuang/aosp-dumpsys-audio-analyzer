const fs = require('fs');
const content = fs.readFileSync('dumpsys_files/music_play.txt', 'utf8');

let threads = [];
let threadRegex = /Output thread 0x([0-9a-fA-F]+).*?Standby:\s*(yes|no).*?(?=Output thread 0x|$)/gs;
let match;
while ((match = threadRegex.exec(content)) !== null) {
    let thread = { id: match[1], standby: match[2], block: match[0] };

    // Tracks
    let tracks = [];
    let trackSection = match[0].match(/Id\s+Active\s+Client.*?(?:\n\s*\n|Effect Chains)/s);
    if (trackSection) {
        let trackLines = trackSection[0].split('\n').slice(1);
        for (let line of trackLines) {
            line = line.trim();
            if (!line || line.includes('Effect Chains')) break;
            let parts = line.split(/\s+/);
            if (parts.length >= 3) {
                tracks.push({ id: parts[0], active: parts[1], client: parts[2] });
            }
        }
    }
    thread.tracks = tracks;

    // Hal frame count
    let halMatch = match[0].match(/HAL frame count:\s*(\d+)/);
    if (halMatch) thread.halFrames = halMatch[1];

    // Frame size
    let sizeMatch = match[0].match(/Processing frame size:\s*(\d+)/);
    if (sizeMatch) thread.frameSize = sizeMatch[1];

    // Sink buffer
    let sinkMatch = match[0].match(/Sink buffer\s*:\s*(0x[0-9a-fA-F]+)/);
    if (sinkMatch) thread.sinkBuffer = sinkMatch[1];
    else {
        // sometimes it's just 'Sink buffer : 0x...' without frames
        let altSink = match[0].match(/Sink buffer\s*:\s*(0x[0-9a-fA-F]+)/);
    }

    threads.push(thread);
}

console.log(JSON.stringify(threads, null, 2));
