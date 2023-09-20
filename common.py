def create_text_chunks(text, max_chunk_size=2250):
    chunks = []
    remaining = text
    while len(remaining) > max_chunk_size:
        # Determine best point `i` to truncate text where i < max_chunk_size
        i = max_chunk_size
        while (remaining[i] != "\n") and (remaining[i-1:i+1] != ". "):
            i -= 1
        # Split remaining text into two — append first to `chunks` and second to `remaining`
        chunks.append(remaining[:i])
        remaining = remaining[i:]
    # Append remaining text to `chunks`
    chunks.append(remaining)
    # Strip all chunks of trailing whitespace and newlines
    for i in range(len(chunks)):
        chunks[i] = chunks[i].strip("\n")
        chunks[i] = chunks[i].strip()
    
    return chunks