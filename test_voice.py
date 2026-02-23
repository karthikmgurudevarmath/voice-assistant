import pyttsx3

try:
    print("Initializing Engine...")
    engine = pyttsx3.init('sapi5')
    
    voices = engine.getProperty('voices')
    print(f"Number of voices: {len(voices)}")
    
    for i, voice in enumerate(voices):
        print(f"Voice {i}: {voice.name} - ID: {voice.id}")

    # Test default
    print("Testing default voice...")
    engine.say("Testing default voice.")
    engine.runAndWait()

    # Test logic used in main app (usually index 1)
    if len(voices) > 1:
        print("Testing Voice 1...")
        engine.setProperty('voice', voices[1].id)
        engine.say("Testing voice one.")
        engine.runAndWait()
    else:
        print("Voice 1 not available.")

    print("Done.")

except Exception as e:
    print(f"Error: {e}")
