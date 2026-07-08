import os, sys, time, glob
os.environ["PATH"] += os.pathsep + r"E:\01.NDRS_WORKON\OpenWork\temp_audio\bin"
os.environ["PATH"] += os.pathsep + r"C:\Users\soraw\AppData\Local\Temp\ffmpeg_extract\ffmpeg-8.1.2-essentials_build\bin"
os.environ["HF_HUB_DISABLE_SYMLINKS_WARNING"] = "1"

from faster_whisper import WhisperModel

output_dir = "E:\\01.NDRS_WORKON\\OpenWork\\ถอดไฟล์เสียง\\BITEC_COLD CHAIN_2026-07-02\\transcript"
os.makedirs(output_dir, exist_ok=True)

print("Loading faster-whisper tiny model (75MB)...")
t0 = time.time()
model = WhisperModel("tiny", device="cpu", compute_type="int8")
print(f"Model loaded in {time.time()-t0:.1f}s")

# Transcribe session 1
src = "E:\\01.NDRS_WORKON\\OpenWork\\temp_audio\\session1.m4a"
print(f"\n=== Session 1 ===")
t0 = time.time()
segments, info = model.transcribe(src, beam_size=1, language="en")
print(f"Detected: {info.language} (p={info.language_probability:.2f})")

txt1, seg1 = [], []
for seg in segments:
    txt1.append(seg.text.strip())
    seg1.append(f"[{seg.start:.2f}s -> {seg.end:.2f}s] {seg.text.strip()}")
t1 = time.time()
print(f"Done in {t1-t0:.1f}s")

# Save session 1
with open(os.path.join(output_dir, "session1_transcript_en.txt"), "w", encoding="utf-8") as f:
    f.write(" ".join(txt1))
with open(os.path.join(output_dir, "session1_segments_en.txt"), "w", encoding="utf-8") as f:
    f.write("\n".join(seg1))
print(f"Session 1: {len(txt1)} segments")

# Transcribe session 2
src2 = "E:\\01.NDRS_WORKON\\OpenWork\\temp_audio\\session2.m4a"
print(f"\n=== Session 2 ===")
t0 = time.time()
segments, info = model.transcribe(src2, beam_size=1, language="en")
print(f"Detected: {info.language} (p={info.language_probability:.2f})")

txt2, seg2 = [], []
for seg in segments:
    txt2.append(seg.text.strip())
    seg2.append(f"[{seg.start:.2f}s -> {seg.end:.2f}s] {seg.text.strip()}")
t2 = time.time()
print(f"Done in {t2-t0:.1f}s")

with open(os.path.join(output_dir, "session2_transcript_en.txt"), "w", encoding="utf-8") as f:
    f.write(" ".join(txt2))
with open(os.path.join(output_dir, "session2_segments_en.txt"), "w", encoding="utf-8") as f:
    f.write("\n".join(seg2))
print(f"Session 2: {len(txt2)} segments")

print(f"\nTotal time: {t2-t1:.1f}s")
print("DONE")
