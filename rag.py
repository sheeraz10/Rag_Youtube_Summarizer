from langchain_huggingface import HuggingFaceEmbeddings
from langchain_chroma import Chroma
from langchain_text_splitters import RecursiveCharacterTextSplitter


# Create embeddings
embeddings = HuggingFaceEmbeddings(
    model_name="sentence-transformers/all-MiniLM-L6-v2"
)


# Create vector database
def create_vectorstore(transcript_text):

    text_splitter = RecursiveCharacterTextSplitter(
        chunk_size=1000,
        chunk_overlap=150
    )

    chunks = text_splitter.split_text(transcript_text)

    vectorstore = Chroma.from_texts(
        texts=chunks,
        embedding=embeddings
    )

    return vectorstore


# Search relevant chunks
def retrieve_relevant_text(vectorstore, question):

    results = vectorstore.similarity_search(
        question,
        k=3
    )

    context = ""

    for result in results:
        context += result.page_content + "\n\n"

    return context