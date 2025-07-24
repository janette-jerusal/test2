Office.onReady(info => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("compareBtn").onclick = async () => {
      try {
        await Excel.run(async (context) => {
          const sheet = context.workbook.worksheets.getActiveWorksheet();
          const range = sheet.getUsedRange();
          range.load("values");
          await context.sync();

          const data = range.values;
          const headers = data[0].map(h => h.toString().toLowerCase());
          const idIndex = headers.indexOf("id");
          const descIndex = headers.indexOf("description");

          if (idIndex === -1 || descIndex === -1) {
            alert("Please ensure your sheet has 'ID' and 'Description' headers.");
            return;
          }

          const entries = data.slice(1).map(row => ({
            id: row[idIndex]?.toString() ?? "",
            description: row[descIndex]?.toString() ?? ""
          })).filter(d => d.id && d.description);

          const results = computeSimilarities(entries, entries);
          populateResults(results);
        });
      } catch (error) {
        console.error(error);
        alert("Error reading from Excel.");
      }
    };
  }
});

// TF-IDF + Cosine Similarity Logic
function computeTFIDF(corpus) {
  const tfidf = [], df = {}, N = corpus.length;
  corpus.forEach((doc, i) => {
    const tokens = tokenize(doc);
    const tf = {};
    tokens.forEach(token => tf[token] = (tf[token] || 0) + 1);
    Object.keys(tf).forEach(term => df[term] = (df[term] || 0) + 1);
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
    const a = vecA[term] || 0, b = vecB[term] || 0;
    dot += a * b;
    normA += a * a;
    normB += b * b;
  });
  return dot / (Math.sqrt(normA) * Math.sqrt(normB) || 1);
}

function computeSimilarities(data1, data2) {
  const corpus = [...data1.map(d => d.description), ...data2.map(d => d.description)];
  const tfidfVectors = computeTFIDF(corpus);
  const results = [];

  for (let i = 0; i < data1.length; i++) {
    const vec1 = tfidfVectors[i], entry1 = data1[i];
    let best = { score: 0, entry2: null };
    for (let j = 0; j < data2.length; j++) {
      const vec2 = tfidfVectors[data1.length + j];
      const score = cosineSimilarity(vec1, vec2);
      if (score > best.score && entry1.id !== data2[j].id) {
        best = { score, entry2: data2[j] };
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
