import os, sys, time
os.environ["PATH"] += os.pathsep + r"E:\01.NDRS_WORKON\OpenWork\temp_audio\bin"
os.environ["PATH"] += os.pathsep + r"C:\Users\soraw\AppData\Local\Temp\ffmpeg_extract\ffmpeg-8.1.2-essentials_build\bin"

from faster_whisper import WhisperModel

audio_file = sys.argv[1]
output_dir = sys.argv[2] if len(sys.argv) > 2 else os.path.dirname(audio_file)
base_name = os.path.splitext(os.path.basename(audio_file))[0]

size = "small"
# Use int8 on CPU for max speed
print(f"Loading faster-whisper model '{size}' (INT8 CPU)...")
t0 = time.time()
model = WhisperModel(size, device="cpu", compute_type="int8")
print(f"Model loaded in {time.time()-t0:.1f}s")

print(f"Transcribing {audio_file}...")
t0 = time.time()
segments, info = model.transcribe(audio_file, beam_size=1, language="en")
print(f"Detected language: {info.language} (p={info.language_probability:.2f})")

txt_lines = []
seg_lines = []
for seg in segments:
    txt_lines.append(seg.text.strip())
    seg_lines.append(f"[{seg.start:.2f}s -> {seg.end:.2f}s] {seg.text.strip()}")

transcript = " ".join(txt_lines)
print(f"Transcription done in {time.time()-t0:.1f}s")

txt_path = os.path.join(output_dir, f"{base_name}_transcript.txt")
with open(txt_path, "w", encoding="utf-8") as f:
    f.write(transcript)
print(f"Saved: {txt_path}")

seg_path = os.path.join(output_dir, f"{base_name}_segments.txt")
with open(seg_path, "w", encoding="utf-8") as f:
    f.write("\n".join(seg_lines))
print(f"Saved: {seg_path}")

print("DONE")
