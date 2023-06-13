# convert-audio-to-text
import speech_recognition as sr
import pyttsx3
import docx
import os


def record_audio():
    r = sr.Recognizer()
    # Mikrofon isimlerini ve indeks numaralarını kontrol et
    print(sr.Microphone.list_microphone_names())
    # Doğru indeks numarasını belirle
    device_index = 1
    with sr.Microphone(device_index=device_index) as source:
        print("Sesinizi kaydedin:")
        try:
            audio = r.listen(source)
            return audio
        except sr.WaitTimeoutError:
            print("Ses algılanmadı.")
        except sr.UnknownValueError:
            print("Ses anlaşılamadı.")
        except sr.RequestError:
            print("Ses hizmetine erişilemiyor.")
    return None





def speech_to_text(audio):
    r = sr.Recognizer()
    try:
        text = r.recognize_google(audio, language="tr-TR")
        return text
    except sr.UnknownValueError:
        print("Ses anlaşılamadı.")
    except sr.RequestError:
        print("Ses hizmetine erişilemiyor.")

def write_to_word(text):
    file_path = "sesli_yazi2.docx"
    if os.path.isfile(file_path):
        doc = docx.Document(file_path)
        paragraphs = doc.paragraphs
        if paragraphs:
            last_paragraph = paragraphs[-1]
            last_paragraph.text += " " + text
        else:
            doc.add_paragraph(text)
    else:
        doc = docx.Document()
        doc.add_paragraph(text)

    doc.save(file_path)
    print("Metin Word belgesine yazıldı.")

def speak_text(text):
    engine = pyttsx3.init()
    engine.say(text)
    engine.runAndWait()


def main():
    print("Sesli yazma programına hoş geldiniz!")
    while True:
        audio = record_audio()
        text = speech_to_text(audio)
        if text:
            print("Algılanan metin: " + text)
            write_to_word(text)
            print("Metin Word belgesine yazıldı.")
            speak_text("Metin Word belgesine yazıldı.")
        else:
            print("Ses anlaşılamadı. Lütfen tekrar deneyin.")
            speak_text("Ses anlaşılamadı. Lütfen tekrar deneyin.")

if __name__ == "__main__":
    main()
