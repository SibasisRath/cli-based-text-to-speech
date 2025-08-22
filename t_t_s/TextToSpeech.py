import gtts
from playsound import playsound
import os
import re

# Optional import for .docx; handled gracefully if missing
try:
    from docx import Document  # pip install python-docx
    DOCX_AVAILABLE = True
except Exception:
    DOCX_AVAILABLE = False


def speak_and_delete(text_to_speak):
    """
    Converts text to speech, plays it, and deletes the temporary file.
    """
    try:
        temp_file = "temp_audio.mp3"
        tts = gtts.gTTS(text_to_speak)
        tts.save(temp_file)
        playsound(temp_file)
        os.remove(temp_file)
    except Exception as e:
        print(f"An error occurred: {e}")
        print("Could not play the audio. Please check your internet connection.")


def split_into_paragraphs(text: str):
    """
    Split text into paragraphs using blank lines as separators.
    - Handles multiple blank lines.
    - Strips leading/trailing whitespace from each paragraph.
    - Filters out empty paragraphs.
    """
    parts = re.split(r'(?:\r?\n){2,}', text.strip())
    paragraphs = [p.strip() for p in parts if p.strip()]
    return paragraphs


def get_folder_and_confirm_overwrite(num_files_expected: int):
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
        print("Invalid input. Please enter 't' for text or 'f' for file.")


def read_text_from_stdin() -> str:
    """
    Read multi-line text from stdin until EOF.
    """
    speak_and_delete("Paste or type your text. Use a blank line between sections to create separate audio files. Finish with Ctrl+D (Linux/Mac) or Ctrl+Z then Enter (Windows)")
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
    - .txt read as UTF-8 by default; if it fails, tries cp1252 as fallback.
    - .docx requires python-docx.
    Returns the extracted text as a single string.
    Raises exceptions with clear messages if something goes wrong.
    """
    if not os.path.isfile(path):
        raise FileNotFoundError(f"File not found: {path}")

    ext = os.path.splitext(path)[1].lower()
    if ext == ".txt":
        # Try UTF-8, fallback to cp1252 for Windows-origin files
        try:
            with open(path, "r", encoding="utf-8") as f:
                return f.read().strip()
        except UnicodeDecodeError:
            with open(path, "r", encoding="cp1252", errors="replace") as f:
                return f.read().strip()

    elif ext == ".docx":
        if not DOCX_AVAILABLE:
            raise RuntimeError(
                "Reading .docx requires 'python-docx'. Install with: pip install python-docx"
            )
        doc = Document(path)
        # Join paragraphs with newlines; consecutive empty paragraphs create blank lines
        paras = []
        for p in doc.paragraphs:
            # Preserve paragraph boundaries
            paras.append(p.text)
        # Normalize to ensure blank lines remain as separators
        return "\n".join(paras).strip()

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
        # File mode
        while True:
            speak_and_delete("Please enter the full path of the text file. It can be a dot t x t or dot doc x file.")
            path = input("Enter path to .txt or .docx file: ").strip().strip('"').strip("'")
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


def main():
    """
    Main function to run the text-to-speech application.
    """
    # Intro
    speak_and_delete("Hi there! Thanks for using the text-to-speech application.")

    # Get user's name
    speak_and_delete("Please enter your name.")
    name = input("Enter your name: ")

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

    # Choose folder and handle overwrite behavior for all outputs
    speak_and_delete("Enter the folder name where I should save the audio files.")
    folder_path = get_folder_and_confirm_overwrite(num_paras)

    # Generate each paragraph as N.mp3
    generated_paths = []
    try:
        for idx, para in enumerate(paragraphs, start=1):
            target = os.path.join(folder_path, f"{idx}.mp3")
            tts = gtts.gTTS(para)
            tts.save(target)
            print(f"Saved: {target}")
            generated_paths.append(target)

        speak_and_delete("All files have been saved successfully.")

        # Ask if should play now
        if ask_playback():
            speak_and_delete("Now playing the audio files one by one.")
            for p in generated_paths:
                playsound(p)
            speak_and_delete("All audio files have been played.")
        else:
            speak_and_delete("Okay, I will not play the audio files now.")

        speak_and_delete("Goodbye!")
    except Exception as e:
        print(f"An error occurred while creating the audio files: {e}")


if __name__ == "__main__":
    main()
