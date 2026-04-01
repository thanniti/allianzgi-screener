# run_pipeline.py
from screener import run_screener
from pitch_generator import build_pitch_slide
import os

def run_full_pipeline(generate_pitch_for_top_n: int = 3):
    funds = run_screener()

    os.makedirs("output", exist_ok=True)

    for fund in funds[:generate_pitch_for_top_n]:
        filename = f"output/{fund['name'].replace(' ', '_')}_Pitch.pptx"
        build_pitch_slide(fund, filename)
        print(f"Generated: {filename}", flush=True)

if __name__ == "__main__":
    run_full_pipeline()