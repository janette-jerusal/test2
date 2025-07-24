Office.onReady(() => {
  document.getElementById("compareBtn").addEventListener("click", async function () {
    const file1 = document.getElementById("file1").files[0];
    const file2 = document.getElementById("file2").files[0];

    if (!file1 || !file2) {
      alert("Please select both Excel files.");
      return;
    }

    const data1 = await readExcel(file1);
    const data2 = await readExcel(file2);

    const results = computeSimilarities(data1, data2);
    populateResults(results);
  });
});

async function readExcel(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = function (e) {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: "array" });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const json = XLSX.utils.sheet_to_json(sheet);
      const cleaned = json.map(row => ({
        id: row.ID?.toString() ?? "",
        description: row.Description?.toString() ?? ""
      }));
      resolve(cleaned);
    };
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}

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

function computeSimilarities(data1, data2) {
  const corpus = [...data1.map(d => d.description), ...data2.map(d => d.description)];
  const tfidfVectors = computeTFIDF(corpus);
  const results = [];

  for (let i = 0; i < data1.length; i++) {
    const vec1 = tfidfVectors[i];
    const entry1 = data1[i];

    let best = { score: 0, entry2: null };
    for (let j = 0; j < data2.length; j++) {
      const vec2 = tfidfVectors[data1.length + j];
      const score = cosineSimilarity(vec1, vec2);
      if (score > best.score) {
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

