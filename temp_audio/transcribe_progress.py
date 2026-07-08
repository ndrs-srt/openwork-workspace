import os, sys, time
os.environ["PATH"] += os.pathsep + r"E:\01.NDRS_WORKON\OpenWork\temp_audio\bin"
os.environ["PATH"] += os.pathsep + r"C:\Users\soraw\AppData\Local\Temp\ffmpeg_extract\ffmpeg-8.1.2-essentials_build\bin"

import whisper

output_dir = "E:\\01.NDRS_WORKON\\OpenWork\\ถอดไฟล์เสียง\\BITEC_COLD CHAIN_2026-07-02\\transcript"
os.makedirs(output_dir, exist_ok=True)

model_name = "tiny"
print(f"[1/6] Loading whisper model '{model_name}'...")
t0 = time.time()
model = whisper.load_model(model_name)
print(f"  Loaded in {time.time()-t0:.1f}s")

# Session 1
src1 = "E:\\01.NDRS_WORKON\\OpenWork\\temp_audio\\session1.m4a"
print(f"\n[2/6] Transcribing Session 1 (11:00-12:00)...")
print(f"  This shows frames processed / total frames")
t0 = time.time()
result1 = model.transcribe(src1, language="en", verbose=True)
t1 = time.time()
print(f"  Done in {t1-t0:.1f}s ({len(result1['segments'])} segments)")

# Save session 1
txt1 = " ".join(seg["text"].strip() for seg in result1["segments"])
seg1 = "\n".join(f"[{seg['start']:.2f}s -> {seg['end']:.2f}s] {seg['text'].strip()}" for seg in result1["segments"])
with open(os.path.join(output_dir, "session1_transcript_en.txt"), "w", encoding="utf-8") as f:
    f.write(txt1)
with open(os.path.join(output_dir, "session1_segments_en.txt"), "w", encoding="utf-8") as f:
    f.write(seg1)
print(f"  Saved session 1")

# Session 2
src2 = "E:\\01.NDRS_WORKON\\OpenWork\\temp_audio\\session2.m4a"
print(f"\n[3/6] Transcribing Session 2 (13:15-14:14)...")
t0 = time.time()
result2 = model.transcribe(src2, language="en", verbose=True)
t2 = time.time()
print(f"  Done in {t2-t0:.1f}s ({len(result2['segments'])} segments)")

txt2 = " ".join(seg["text"].strip() for seg in result2["segments"])
seg2 = "\n".join(f"[{seg['start']:.2f}s -> {seg['end']:.2f}s] {seg['text'].strip()}" for seg in result2["segments"])
with open(os.path.join(output_dir, "session2_transcript_en.txt"), "w", encoding="utf-8") as f:
    f.write(txt2)
with open(os.path.join(output_dir, "session2_segments_en.txt"), "w", encoding="utf-8") as f:
    f.write(seg2)
print(f"  Saved session 2")

total = t2 - t1 + (t1 - time.time() + (t2 - t0))
print(f"\n[4/6] Transcription complete!")
print(f"  Session 1: {t1-time.time()+t2-t0:.1f}s")
print(f"  Session 2: {t2-t0:.1f}s")
print(f"  Files saved to: {output_dir}")
print("Ready for Excel summary.")
