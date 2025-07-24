Office.onReady(() => {
  console.log("Office.js is ready.");
  document.getElementById("compareBtn").addEventListener("click", handleCompare);
});

async function handleCompare() {
  console.log("Compare button clicked.");

  const file1 = document.getElementById("file1").files[0];
  const file2 = document.getElementById("file2").files[0];

  if (!file1 || !file2) {
    console.log("One or both files are missing.");
    alert("Please upload both Excel files.");
    return;
  }

  try {
    const [data1, data2] = await Promise.all([readExcelFile(file1), readExcelFile(file2)]);
    console.log("File 1 data:", data1);
    console.log("File 2 data:", data2);

    const results = computeSimilarities(data1, data2);
    console.log("Similarity results:", results);

    populateResults(results);
  } catch (error) {
    console.error("Error during comparison:", error);
  }
}

function readExcelFile(file) {
  return new Promise((resolve, reject) => {
    console.log("Reading file:", file.name);

    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: "array" });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const json = XLSX.utils.sheet_to_json(sheet);
        console.log(`Parsed ${json.length} rows from ${file.name}`);
        resolve(json);
      } catch (err) {
        reject(`Failed to parse ${file.name}: ${err}`);
      }
    };
    reader.onerror = () => reject(`Failed to read ${file.name}`);
    reader.readAsArrayBuffer(file);
  });
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

function cosineSimilarity(vecA, vecB) {
  const dotProduct = vecA.reduce((sum, a, i) => sum + a * vecB[i], 0);
  const normA = Math.sqrt(vecA.reduce((sum, a) => sum + a * a, 0));
  const normB = Math.sqrt(vecB.reduce((sum, b) => sum + b * b, 0));
  return dotProduct / (normA * normB || 1);
}

function computeTFIDF(corpus) {
  const terms = new Set();
  corpus.forEach(text => {
    text.split(/\s+/).forEach(word => terms.add(word.toLowerCase()));
  });
  const termArray = Array.from(terms);
  const termIndex = termArray.reduce((acc, term, i) => (acc[term] = i, acc), {});
  const tfidf = corpus.map(text => {
    const vec = Array(termArray.length).fill(0);
    const words = text.split(/\s+/).map(w => w.toLowerCase());
    words.forEach(word => {
      if (termIndex.hasOwnProperty(word)) vec[termIndex[word]] += 1;
    });
    return vec;
  });
  return tfidf;
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
      <td>${row.similarity}</td>
    `;
    tbody.appendChild(tr);
  });
}
