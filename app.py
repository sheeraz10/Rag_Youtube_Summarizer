import os
import time
from urllib.parse import urlparse, parse_qs

import streamlit as st
from dotenv import load_dotenv
from youtube_transcript_api import YouTubeTranscriptApi

from google import genai
from google.genai import errors

from rag import create_vectorstore, retrieve_relevant_text
from presentation import create_presentation

# --------------------------------
# Setup
# --------------------------------

load_dotenv()

client = genai.Client(
    api_key=os.getenv("GEMINI_API_KEY") or os.getenv("GOOGLE_API_KEY")
)

# Tried in order. If the first is overloaded (503), the next one is used.
MODELS = [
    "gemini-3.6-flash",
    "gemini-3.5-flash-lite",
]

TRANSCRIPT_LANGUAGES = ["en", "hi", "ur"]

SUMMARY_PROMPT = """You summarize YouTube transcripts.

Rules:
- Use only what's in the transcript. Never add outside information.
- Cut greetings, jokes, filler, and repetition.
- Keep the main ideas, explanations, examples, and conclusions.
- Short bullets, no long paragraphs.
- Be concise, but don't drop anything important.
- Briefly define technical terms in plain language.
- Skip any section the transcript doesn't support.

Format the answer using Markdown:

## Summary

### Main Topic
- ...

### Important Concepts
- Term: one-line explanation

### Key Points
- ...

### Steps / Process
1. ...
2. ...
3. ...

### Final Takeaway
- ...
"""


# --------------------------------
# Gemini helper with retry + fallback
# --------------------------------

def call_gemini(contents, retries_per_model=3):
    """Call Gemini, retrying on overload (503/500/429) and falling back
    to other models. Returns text, or None if everything failed."""

    failures = {}  # model -> last error, so the real cause isn't hidden

    for model in MODELS:
        for attempt in range(retries_per_model):
            try:
                response = client.models.generate_content(
                    model=model,
                    contents=contents,
                )
                return response.text

            except errors.ServerError as e:
                # 5xx: overloaded / internal error -> wait and retry
                failures[model] = e
                if attempt < retries_per_model - 1:
                    time.sleep(2 ** attempt * 2)  # 2s, 4s

            except errors.ClientError as e:
                failures[model] = e
                if getattr(e, "code", None) == 429:
                    # rate limited -> wait and retry
                    if attempt < retries_per_model - 1:
                        time.sleep(2 ** attempt * 3)
                    continue
                # 400/401/403/404 won't fix themselves on retry;
                # move on to the next model
                break

            except Exception as e:
                failures[model] = e
                break

    st.error("All Gemini models failed:")
    for model, err in failures.items():
        st.write(f"**{model}**: {err}")
    return None


# --------------------------------
# Extract video ID
# --------------------------------

def get_video_id(youtube_url):
    parsed_url = urlparse(youtube_url)
    host = parsed_url.netloc

    if "youtube.com" in host:
        if parsed_url.path == "/watch":
            return parse_qs(parsed_url.query).get("v", [None])[0]
        # /shorts/ID, /embed/ID, /live/ID
        parts = [p for p in parsed_url.path.split("/") if p]
        if len(parts) >= 2 and parts[0] in ("shorts", "embed", "live"):
            return parts[1]
        return None

    if "youtu.be" in host:
        return parsed_url.path.strip("/") or None

    return None


# --------------------------------
# Extract transcript
# --------------------------------

def extract_transcript_details(youtube_url):
    video_id = get_video_id(youtube_url)

    if not video_id:
        st.error("Invalid YouTube URL")
        return None

    try:
        api = YouTubeTranscriptApi()
        transcript = api.fetch(video_id, languages=TRANSCRIPT_LANGUAGES)
        return " ".join(entry.text for entry in transcript)

    except Exception as e:
        st.error(f"Could not get transcript: {e}")
        return None


# --------------------------------
# Streamlit UI
# --------------------------------

st.title("YouTube Transcript to Detailed Notes Converter")

youtube_link = st.text_input("Enter YouTube URL")

if youtube_link:
    thumb_id = get_video_id(youtube_link)
    if thumb_id:
        st.image(
            f"https://img.youtube.com/vi/{thumb_id}/0.jpg",
            caption="Video Thumbnail",
            use_container_width=True,
        )


# --------------------------------
# Generate notes
# --------------------------------

if st.button("Generate Notes"):

    if not youtube_link:
        st.warning("Please enter a YouTube URL.")

    else:
        # Clear results from any previous video
        for key in ("summary", "ppt_bytes", "vectorstore"):
            st.session_state.pop(key, None)

        with st.spinner("Fetching transcript..."):
            transcript_text = extract_transcript_details(youtube_link)

        if transcript_text:

            with st.spinner("Building search index..."):
                try:
                    st.session_state.vectorstore = create_vectorstore(
                        transcript_text
                    )
                except Exception as e:
                    st.error(f"Could not build the vector store: {e}")

            with st.spinner("Generating notes..."):
                summary = call_gemini(
                    f"{SUMMARY_PROMPT}\n\nTranscript:\n{transcript_text}"
                )

            if summary:
                st.session_state.summary = summary

                with st.spinner("Creating PowerPoint..."):
                    try:
                        ppt_file = create_presentation(
                            "YouTube Video Summary", summary
                        )
                        with open(ppt_file, "rb") as f:
                            st.session_state.ppt_bytes = f.read()
                    except Exception as e:
                        st.error(f"Could not create the PowerPoint: {e}")
            else:
                st.warning("Summary could not be generated. Please try again.")


# --------------------------------
# Show saved results (survives reruns caused by download buttons)
# --------------------------------

if "summary" in st.session_state:
    st.markdown(st.session_state.summary)

    st.download_button(
        "Download Notes",
        st.session_state.summary,
        file_name="youtube_notes.md",
        mime="text/markdown",
    )

    if "ppt_bytes" in st.session_state:
        st.download_button(
            "Download PowerPoint",
            st.session_state.ppt_bytes,
            file_name="youtube_summary.pptx",
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "presentationml.presentation"
            ),
        )


# --------------------------------
# Ask questions about the video
# --------------------------------

st.divider()
st.header("Ask Questions About This Video")

question = st.text_input("Ask a question about the video")

if st.button("Ask Question"):

    if "vectorstore" not in st.session_state:
        st.warning("First generate notes for a YouTube video.")

    elif not question:
        st.warning("Please enter a question.")

    else:
        try:
            context = retrieve_relevant_text(
                st.session_state.vectorstore, question
            )
        except Exception as e:
            st.error(f"Could not search the transcript: {e}")
            context = None

        if context:
            rag_prompt = f"""
Answer the user's question using ONLY the context below.

If the answer is not available in the context,
say: "This information is not available in the video."

Context:
{context}

Question:
{question}
"""
            with st.spinner("Thinking..."):
                answer = call_gemini(rag_prompt)

            if answer:
                st.subheader("Answer")
                st.write(answer)