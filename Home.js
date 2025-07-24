document.getElementById("compareBtn").addEventListener("click", async function () {
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const usedRange = sheet.getUsedRange();
      usedRange.load("values");
      await context.sync();

      const rows = usedRange.values;
      const headers = rows[0];
      const idIndex = headers.findIndex((h) => h.toLowerCase().includes("id"));
      const descIndex = headers.findIndex((h) => h.toLowerCase().includes("desc"));

      if (idIndex === -1 || descIndex === -1) {
        alert("Couldn't find columns for ID and Description.");
        return;
      }

      const data = rows.slice(1) // skip header
        .map(row => ({
          id: row[idIndex]?.toString() ?? "",
          description: row[descIndex]?.toString() ?? ""
        }))
        .filter(d => d.id && d.description);

      if (data.length < 2) {
        alert("Need at least 2 user stories to compare.");
        return;
      }

      const results = computeSimilarities(data);
      populateResults(results);
    });
  } catch (error) {
    console.error(error);
    alert("Error reading from Excel. Make sure your file is open and contains ID and Description columns.");
  }
});

// TF-IDF + Cosine Similarity Logic
function computeTFIDF(corpus) {
  const tfidf = [];
  const df = {};
  const N = corpus.length;

  corpus.forEach((doc, i) => {
    const tokens = tokenize(doc);
    const tf = {};
    tokens.forEach(token => {
      tf[token] = (tf[token] || 0) + 1;
    });
    Object.keys(tf).forEach(term => {
      df[term] = (df[term] || 0) + 1;
    });
    tfidf[i] = tf;
  });

  return tfidf.map(tf => {
    const vector = {};
    Object.keys(tf).forEach(term => {
      const tfVal = tf[term];
      const idf = Math.log(N / (1 + df[term]));
      vector[term] = tfVal * idf;
    });
    return vector;
  });
}

function tokenize(text) {
  return text.toLowerCase().replace(/[^\w\s]/g, "").split(/\s+/).filter(Boolean);
}

function cosineSimilarity(vecA, vecB) {
  const terms = new Set([...Object.keys(vecA), ...Object.keys(vecB)]);
  let dot = 0, normA = 0, normB = 0;
  terms.forEach(term => {
    const a = vecA[term] || 0;
    const b = vecB[term] || 0;
    dot += a * b;
    normA += a * a;
    normB += b * b;
  });
  return dot / (Math.sqrt(normA) * Math.sqrt(normB) || 1);
}

function computeSimilarities(data) {
  const corpus = data.map(d => d.description);
  const tfidfVectors = computeTFIDF(corpus);
  const results = [];

  for (let i = 0; i < data.length; i++) {
    const vec1 = tfidfVectors[i];
    const entry1 = data[i];

    let best = { score: 0, entry2: null };
    for (let j = 0; j < data.length; j++) {
      if (i === j) continue;
      const vec2 = tfidfVectors[j];
      const score = cosineSimilarity(vec1, vec2);
      if (score > best.score) {
        best = { score, entry2: data[j] };
      }
    }

    if (best.entry2) {
      results.push({
        id1: entry1.id,
        desc1: entry1.description,
        id2: best.entry2.id,
        desc2: best.entry2.description,
        similarity: (best.score * 100).toFixed(2)
      });
    }
  }

  return results;
}

function populateResults(results) {
  const tbody = document.querySelector("#resultTable tbody");
  tbody.innerHTML = "";

  results.forEach(row => {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${row.id1}</td>
      <td>${row.desc1}</td>
      <td>${row.id2}</td>
      <td>${row.desc2}</td>
      <td>${row.similarity}%</td>
    `;
    tbody.appendChild(tr);
  });
}

