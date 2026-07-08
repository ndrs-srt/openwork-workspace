import os, sys, time, json, subprocess
os.environ["PATH"] += os.pathsep + r"E:\01.NDRS_WORKON\OpenWork\temp_audio\bin"
os.environ["PATH"] += os.pathsep + r"C:\Users\soraw\AppData\Local\Temp\ffmpeg_extract\ffmpeg-8.1.2-essentials_build\bin"
os.environ["HF_HUB_DISABLE_SYMLINKS_WARNING"] = "1"

import faster_whisper

output_dir = "E:\\01.NDRS_WORKON\\OpenWork\\ถอดไฟล์เสียง\\BITEC_COLD CHAIN_2026-07-02\\transcript"
os.makedirs(output_dir, exist_ok=True)

def get_duration(path):
    r = subprocess.run([
        "ffprobe", "-v", "error", "-show_entries", "format=duration",
        "-of", "default=noprint_wrappers=1:nokey=1", path
    ], capture_output=True, text=True)
    return float(r.stdout.strip())

def transcribe_file(src, label):
    dur = get_duration(src)
    print(f"\n{'='*60}")
    print(f"{label}")
    print(f"Duration: {dur/60:.1f} min ({dur:.0f}s)")
    print(f"Audo: {os.path.basename(src)}")
    print(f"{'='*60}")

    model = faster_whisper.WhisperModel("tiny", device="cpu", compute_type="int8")

    t0 = time.time()
    segments, info = model.transcribe(src, beam_size=1, language="en")
    print(f"Language: {info.language} (p={info.language_probability:.2f})")

    txt_parts = []
    seg_lines = []
    last_pct = 0
    count = 0

    for seg in segments:
        txt_parts.append(seg.text.strip())
        seg_lines.append(f"[{seg.start:.2f}s -> {seg.end:.2f}s] {seg.text.strip()}")
        count += 1

        pct = min(int((seg.end / dur) * 100), 100)
        if pct > last_pct or count % 20 == 0:
            elapsed = time.time() - t0
            speed = seg.end / elapsed if elapsed > 0 else 0
            eta = (dur - seg.end) / speed if speed > 0 else 0
            bar = "█" * (pct // 5) + "░" * (20 - pct // 5)
            sys.stdout.write(f"\r  [{bar}] {pct}% | {count} segments | {seg.end:.0f}s/{dur:.0f}s | speed: {speed:.1f}x | ETA: {eta:.0f}s  ")
            sys.stdout.flush()
            last_pct = pct

    elapsed = time.time() - t0
    print(f"\n  Done in {elapsed:.1f}s ({elapsed/60:.1f}min) at {dur/elapsed:.1f}x realtime")
    return " ".join(txt_parts), seg_lines

# Session 1
txt1, seg1 = transcribe_file(
    "E:\\01.NDRS_WORKON\\OpenWork\\temp_audio\\session1.m4a",
    "SESSION 1: 11:00-12:00 - GDP Standards & Temperature Monitoring"
)

with open(os.path.join(output_dir, "session1_transcript_en.txt"), "w", encoding="utf-8") as f:
    f.write(txt1)
with open(os.path.join(output_dir, "session1_segments_en.txt"), "w", encoding="utf-8") as f:
    f.write("\n".join(seg1))
print(f"  Saved {len(seg1)} segments")

# Session 2
txt2, seg2 = transcribe_file(
    "E:\\01.NDRS_WORKON\\OpenWork\\temp_audio\\session2.m4a",
    "SESSION 2: 13:15-14:14 - Cold Chain Logistics & Risk Management"
)

with open(os.path.join(output_dir, "session2_transcript_en.txt"), "w", encoding="utf-8") as f:
    f.write(txt2)
with open(os.path.join(output_dir, "session2_segments_en.txt"), "w", encoding="utf-8") as f:
    f.write("\n".join(seg2))
print(f"  Saved {len(seg2)} segments")

print(f"\n{'='*60}")
print(f"ALL DONE! Transcripts saved to:")
print(f"  {output_dir}")
print(f"{'='*60}")
