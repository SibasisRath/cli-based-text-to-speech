import os
import re
import time
from io import BytesIO
from typing import List, Tuple
from concurrent.futures import ThreadPoolExecutor, as_completed

import gtts
from playsound import playsound

# Optional import for .docx support
try:
    from docx import Document  # pip install python-docx
    DOCX_AVAILABLE = True
except Exception:
    DOCX_AVAILABLE = False


# =========================
# Voice prompt helper
# =========================
def speak_and_delete(text_to_speak: str):
    """
    Converts text to speech, plays it, and deletes the temporary file.
    Keep this simple and blocking to avoid race conditions.
    """
    try:
        temp_file = "temp_audio.mp3"
        tts = gtts.gTTS(text_to_speak)
        tts.save(temp_file)
        # small guard in case FS is slow
        time.sleep(0.05)
        playsound(temp_file)
        # guard to avoid racing the OS
        time.sleep(0.05)
        if os.path.exists(temp_file):
            os.remove(temp_file)
    except Exception as e:
        print(f"[speak] An error occurred: {e}")
        print("Could not play the audio. Please check your internet connection.")


# =========================
# Text handling
# =========================
def split_into_paragraphs(text: str) -> List[str]:
    """
    Split text into paragraphs using blank lines as separators.
    Handles multiple blank lines and trims whitespace.
    """
    parts = re.split(r'(?:\r?\n){2,}', text.strip())
    paragraphs = [p.strip() for p in parts if p.strip()]
    return paragraphs


def choose_input_mode() -> str:
    """
    Ask user to choose between direct text (t) or file (f).
    Returns 't' or 'f'.
    """
    while True:
        speak_and_delete("Would you like to enter text directly, or use a text file? Enter t for text or f for file.")
        mode = input("Enter text directly or use a text file? (t/f): ").strip().lower()
        if mode in ('t', 'f'):
            return mode
        print("Invalid input. Please enter 't' or 'f'.")


def read_text_from_stdin() -> str:
    """
    Read multi-line text from stdin until EOF.
    """
    speak_and_delete("Paste or type your text. Use a blank line between sections to create separate audio files.")
    print("Enter your text. Use blank lines to separate paragraphs.")
    print("Finish with Ctrl+D (Linux/Mac) or Ctrl+Z then Enter (Windows):")
    print("(Start typing below)")
    lines = []
    try:
        while True:
            line = input()
            lines.append(line)
    except EOFError:
        pass
    return "\n".join(lines).strip()


def read_text_from_file(path: str) -> str:
    """
    Read text from a .txt or .docx file.
    - .txt read as UTF-8; fallback to cp1252 on decode errors.
    - .docx requires python-docx.
    Returns the extracted text.
    """
    path = path.strip().strip('"').strip("'")
    if not os.path.isfile(path):
        raise FileNotFoundError(f"File not found: {path}")

    ext = os.path.splitext(path)[1].lower()
    if ext == ".txt":
        try:
            with open(path, "r", encoding="utf-8") as f:
                return f.read().strip()
        except UnicodeDecodeError:
            with open(path, "r", encoding="cp1252", errors="replace") as f:
                return f.read().strip()

    elif ext == ".docx":
        if not DOCX_AVAILABLE:
            raise RuntimeError("Reading .docx requires 'python-docx'. Install with: pip install python-docx")
        doc = Document(path)
        # Preserve paragraph boundaries; Word paragraphs map naturally here
        return "\n".join(p.text for p in doc.paragraphs).strip()

    else:
        raise ValueError("Unsupported file type. Please provide a .txt or .docx file.")


def get_text_via_choice() -> str:
    """
    Orchestrate input mode selection and return text.
    """
    mode = choose_input_mode()
    if mode == 't':
        return read_text_from_stdin()
    else:
        while True:
            speak_and_delete("Please enter the full path of the text file. It can be a dot t x t or dot doc x file.")
            path = input("Enter path to .txt or .docx file: ").strip()
            try:
                text = read_text_from_file(path)
                if not text:
                    print("The file appears to be empty. Please choose another file.")
                    speak_and_delete("The file appears to be empty. Please choose another file.")
                    continue
                return text
            except Exception as e:
                print(f"Could not read the file: {e}")
                speak_and_delete("Could not read the file. Please try again.")


# =========================
# Output folder and overwrite handling
# =========================
def get_folder_and_confirm_overwrite(num_files_expected: int) -> str:
    """
    Ask for a folder name, create it if needed.
    Target filenames will be 1.mp3 .. N.mp3 inside that folder.
    If any of those files already exist, ask once whether to overwrite all.
    If 'n', re-ask for folder.
    Returns absolute folder path.
    """
    while True:
        folder_name = input("Enter the folder name: ").strip()
        if not folder_name:
            speak_and_delete("Folder name cannot be empty. Please try again.")
            print("Folder name cannot be empty. Please try again.")
            continue

        try:
            os.makedirs(folder_name, exist_ok=True)
        except Exception as e:
            print(f"Could not create or access the folder '{folder_name}': {e}")
            speak_and_delete("Could not create or access the folder. Please try a different name.")
            continue

        folder_abs = os.path.abspath(folder_name)
        targets = [os.path.join(folder_abs, f"{i}.mp3") for i in range(1, num_files_expected + 1)]
        conflicts = [p for p in targets if os.path.exists(p)]

        if conflicts:
            while True:
                speak_and_delete(f"{len(conflicts)} files already exist in the folder. "
                                 f"Do you want to overwrite them? Enter y for yes or n for no.")
                choice = input(f"{len(conflicts)} target files already exist. Overwrite all? (y/n): ").strip().lower()
                if choice == 'y':
                    for p in conflicts:
                        try:
                            os.remove(p)
                        except Exception as e:
                            print(f"Warning: Could not remove existing file '{p}': {e}")
                    return folder_abs
                elif choice == 'n':
                    speak_and_delete("Please enter a different folder name.")
                    print("Please enter a different folder name.")
                    break
                else:
                    speak_and_delete("Invalid input. Please enter y for yes or n for no.")
                    print("Invalid input. Please enter 'y' or 'n'.")
        else:
            return folder_abs


def ask_playback() -> bool:
    """
    Ask the user whether to play the generated audio files now.
    Returns True for yes, False for no. Re-prompts on invalid input.
    """
    while True:
        speak_and_delete("Do you want me to play the audio files now? Please enter y for yes or n for no.")
        ans = input("Play the audio files now? (y/n): ").strip().lower()
        if ans in ('y', 'n'):
            return ans == 'y'
        print("Invalid input. Please enter 'y' or 'n'.")


# =========================
# Parallel TTS generation
# =========================
def tts_bytes_with_retry(text: str, retries: int = 3, base_delay: float = 0.4) -> bytes:
    """
    Convert text to MP3 bytes using gTTS with simple retry and exponential backoff.
    Raises the last exception if all attempts fail.
    """
    last_err = None
    for attempt in range(1, retries + 1):
        try:
            buf = BytesIO()
            tts = gtts.gTTS(text)
            tts.write_to_fp(buf)
            buf.seek(0)
            return buf.read()
        except Exception as e:
            last_err = e
            # backoff: 0.4, 0.8, 1.6, ...
            time.sleep(base_delay * (2 ** (attempt - 1)))
    raise last_err


def generate_all_parallel(
    paragraphs: List[str],
    out_folder: str,
    max_workers: int = 6,
    rate_limit_delay: float = 0.0
) -> Tuple[List[str], List[Tuple[int, Exception]]]:
    """
    Convert paragraphs to MP3 files in parallel.
    - paragraphs: list of paragraph strings
    - out_folder: destination directory
    - max_workers: thread pool size
    - rate_limit_delay: optional small sleep before each task to avoid burst

    Returns:
      (success_paths, failures) where:
        success_paths: list of absolute file paths written (sorted 1..N)
        failures: list of (index, exception) for failed conversions
    """
    os.makedirs(out_folder, exist_ok=True)
    success_paths: List[str] = []
    failures: List[Tuple[int, Exception]] = []

    def task(idx: int, text: str) -> str:
        # optional rate limit to avoid hammering the service
        if rate_limit_delay > 0:
            time.sleep(rate_limit_delay)
        data = tts_bytes_with_retry(text, retries=3, base_delay=0.4)
        path = os.path.join(out_folder, f"{idx}.mp3")
        with open(path, "wb") as f:
            f.write(data)
        return path

    # Submit tasks
    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        futures = {executor.submit(task, idx, para): idx for idx, para in enumerate(paragraphs, start=1)}
        for fut in as_completed(futures):
            idx = futures[fut]
            try:
                path = fut.result()
                abs_path = os.path.abspath(path)
                success_paths.append(abs_path)
                print(f"Saved: {abs_path}")
            except Exception as e:
                failures.append((idx, e))
                print(f"Error generating {idx}.mp3: {e}")

    # Keep output order stable (1..N)
    success_paths.sort(key=lambda p: int(os.path.splitext(os.path.basename(p))[0]))
    # Sort failures by index for stable reporting
    failures.sort(key=lambda x: x)
    return success_paths, failures


# =========================
# Main app flow
# =========================
def main():
    # Intro
    speak_and_delete("Hi there! Thanks for using the text to speech application.")

    # Get user's name
    speak_and_delete("Please enter your name.")
    name = input("Enter your name: ").strip()

    # Choose input method and obtain text
    full_text = get_text_via_choice()
    if not full_text:
        speak_and_delete("No text provided. Exiting.")
        print("No text provided. Exiting.")
        return

    # Split into paragraphs
    paragraphs = split_into_paragraphs(full_text)
    num_paras = len(paragraphs)
    speak_and_delete(f"Detected {num_paras} paragraph{'s' if num_paras != 1 else ''}.")
    print(f"Detected {num_paras} paragraph(s).")

    if num_paras == 0:
        speak_and_delete("No valid paragraphs found. Exiting.")
        print("No valid paragraphs found. Exiting.")
        return

    # Choose folder and handle overwrite behavior for all outputs
    speak_and_delete("Enter the folder name where I should save the audio files.")
    folder_path = get_folder_and_confirm_overwrite(num_paras)

    # Performance settings (tunable)
    # - max_workers: increase to speed up, but be mindful of rate limits
    # - rate_limit_delay: small delay before each task to avoid bursts (0.0 for max speed)
    max_workers = min(8, max(2, os.cpu_count() or 4))  # reasonable default bound
    rate_limit_delay = 0.0  # set to 0.1 or 0.2 if you encounter throttling

    # Parallel generation
    generated_paths, failures = generate_all_parallel(
        paragraphs,
        folder_path,
        max_workers=max_workers,
        rate_limit_delay=rate_limit_delay
    )

    # Report results
    if failures and generated_paths:
        speak_and_delete("Some files could not be generated. Please check the console for details.")
    elif failures and not generated_paths:
        speak_and_delete("No files could be generated due to errors.")
    else:
        speak_and_delete("All files have been saved successfully.")

    # List failures (if any)
    if failures:
        print("Failures:")
        for idx, err in failures:
            print(f"- {idx}.mp3 -> {err}")

    # Playback prompt
    if generated_paths:
        if ask_playback():
            speak_and_delete("Now playing the audio files one by one.")
            # Play in order 1..N
            for p in generated_paths:
                try:
                    playsound(p)
                except Exception as e:
                    print(f"Playback error for {p}: {e}")
            speak_and_delete("All available audio files have been played.")
        else:
            speak_and_delete("Okay, I will not play the audio files now.")

    speak_and_delete("Goodbye!")


if __name__ == "__main__":
    main()
