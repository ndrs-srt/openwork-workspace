import os, sys, time, glob
os.environ["PATH"] += os.pathsep + r"E:\01.NDRS_WORKON\OpenWork\temp_audio\bin"
os.environ["PATH"] += os.pathsep + r"C:\Users\soraw\AppData\Local\Temp\ffmpeg_extract\ffmpeg-8.1.2-essentials_build\bin"
os.environ["HF_HUB_DISABLE_SYMLINKS_WARNING"] = "1"

from faster_whisper import WhisperModel

chunk_dir = "E:\\01.NDRS_WORKON\\OpenWork\\temp_audio\\chunks"
output_dir = "E:\\01.NDRS_WORKON\\OpenWork\\ถอดไฟล์เสียง\\BITEC_COLD CHAIN_2026-07-02\\transcript"
os.makedirs(output_dir, exist_ok=True)

chunks = sorted(glob.glob(os.path.join(chunk_dir, "chunk_*.m4a")))
print(f"Found {len(chunks)} chunks")

print("Loading tiny model (fast)...")
t0 = time.time()
model = WhisperModel("tiny", device="cpu", compute_type="int8")
print(f"Model loaded in {time.time()-t0:.1f}s")

all_segments = []
for i, chunk_path in enumerate(chunks):
    print(f"\n--- Chunk {i+1}/{len(chunks)}: {os.path.basename(chunk_path)} ---")
    t0 = time.time()
    segments, info = model.transcribe(chunk_path, beam_size=1, language="en")
    chunk_text = []
    for seg in segments:
        line = f"[{seg.start:.2f}s -> {seg.end:.2f}s] {seg.text.strip()}"
        chunk_text.append(seg.text.strip())
        all_segments.append((seg.start + (i*300), seg.end + (i*300), seg.text.strip()))
    print(f"  Done in {time.time()-t0:.1f}s")

# Sort by adjusted start time and write combined
all_segments.sort(key=lambda x: x[0])
combined_text = " ".join(seg[2] for seg in all_segments)
segments_text = "\n".join(f"[{s:.2f}s -> {e:.2f}s] {t}" for s, e, t in all_segments)

txt_path = os.path.join(output_dir, "session1_transcript.txt")
with open(txt_path, "w", encoding="utf-8") as f:
    f.write(combined_text)

seg_path = os.path.join(output_dir, "session1_segments.txt")
with open(seg_path, "w", encoding="utf-8") as f:
    f.write(segments_text)

print(f"\nSaved: {txt_path}")
print(f"Saved: {seg_path}")
print(f"Total segments: {len(all_segments)}")
print("DONE")
