import { useMemo, useState } from "react";
import * as XLSX from "xlsx";
import "./App.css";

function normalizeTitle(title) {
  return title.trim().replace(/\s+/g, " ").toLowerCase();
}

function parseTracks(text) {
  const tracks = [];

  // Cho phép:
  // TRACK 31: Title
  // **TRACK 31: Title
  // #### TRACK 31: Title
  // ### **TRACK 31: Title
  const regex = /(?:^|\n)\s*(?:[#>*`_\s]*)TRACK\s+(\d+)\s*:\s*(.+?)(?=\r?\n|$)/gi;

  let match;

  while ((match = regex.exec(text)) !== null) {
    const trackNumber = Number(match[1]);
    const title = match[2]
      .replace(/[`*_]+$/g, "")
      .trim();

    if (title) {
      tracks.push({
        trackNumber,
        title,
      });
    }
  }

  return tracks;
}

function analyzeDocuments(documents) {
  const allTitleOccurrences = new Map();

  // Đếm title trong toàn bộ documents
  documents.forEach((doc) => {
    doc.tracks.forEach((track) => {
      const key = normalizeTitle(track.title);

      if (!allTitleOccurrences.has(key)) {
        allTitleOccurrences.set(key, []);
      }

      allTitleOccurrences.get(key).push({
        documentId: doc.id,
        documentName: doc.name,
      });
    });
  });

  return documents.map((doc) => {
    const localCount = new Map();

    doc.tracks.forEach((track) => {
      const key = normalizeTitle(track.title);
      localCount.set(key, (localCount.get(key) || 0) + 1);
    });

    return {
      ...doc,
      tracks: doc.tracks.map((track) => {
        const key = normalizeTitle(track.title);

        const duplicateInDocument =
          localCount.get(key) > 1;

        const occurrences =
          allTitleOccurrences.get(key) || [];

        const documentIds = new Set(
          occurrences.map((item) => item.documentId)
        );

        const duplicateAcrossDocuments =
          documentIds.size > 1;

        return {
          ...track,
          duplicateInDocument,
          duplicateAcrossDocuments,
        };
      }),
    };
  });
}

function App() {
  const [documents, setDocuments] = useState([]);
  const [showModal, setShowModal] = useState(false);
  const [documentName, setDocumentName] = useState("");
  const [documentText, setDocumentText] = useState("");
  const [search, setSearch] = useState("");
  const [filter, setFilter] = useState("all");

  const analyzedDocuments = useMemo(
    () => analyzeDocuments(documents),
    [documents]
  );

  const totalTracks = analyzedDocuments.reduce(
    (sum, doc) => sum + doc.tracks.length,
    0
  );

  const duplicateTitles = useMemo(() => {
    const map = new Map();

    analyzedDocuments.forEach((doc) => {
      doc.tracks.forEach((track) => {
        const key = normalizeTitle(track.title);

        if (!map.has(key)) {
          map.set(key, {
            title: track.title,
            documents: new Set(),
            count: 0,
            localDuplicate: false,
            crossDuplicate: false,
          });
        }

        const item = map.get(key);

        item.documents.add(doc.name);
        item.count += 1;

        if (track.duplicateInDocument) {
          item.localDuplicate = true;
        }

        if (track.duplicateAcrossDocuments) {
          item.crossDuplicate = true;
        }
      });
    });

    return [...map.values()].filter(
      (item) => item.count > 1
    );
  }, [analyzedDocuments]);

  function addDocument() {
    if (!documentText.trim()) {
      alert("Hãy paste nội dung document.");
      return;
    }

    const tracks = parseTracks(documentText);

    if (tracks.length === 0) {
      alert("Không tìm thấy TRACK nào.");
      return;
    }

    const name =
      documentName.trim() ||
      `Document ${documents.length + 1}`;

    const newDocument = {
      id: Date.now(),
      name,
      tracks,
    };

    setDocuments((prev) => [...prev, newDocument]);

    setDocumentName("");
    setDocumentText("");
    setShowModal(false);
  }

  function removeDocument(id) {
    setDocuments((prev) =>
      prev.filter((doc) => doc.id !== id)
    );
  }

  function exportExcel() {
    if (analyzedDocuments.length === 0) {
      alert("Chưa có document nào.");
      return;
    }

    const trackRows = [];

    analyzedDocuments.forEach((doc) => {
      doc.tracks.forEach((track) => {
        let status = "";

        if (
          track.duplicateInDocument &&
          track.duplicateAcrossDocuments
        ) {
          status = "🔴🟠";
        } else if (track.duplicateInDocument) {
          status = "🔴";
        } else if (track.duplicateAcrossDocuments) {
          status = "🟠";
        }

        trackRows.push({
          Document: doc.name,
          "Track No.": track.trackNumber,
          Title: track.title,
          Status: status,
        });
      });
    });

    const summaryRows = duplicateTitles.map((item) => {
      let type = "";

      if (
        item.localDuplicate &&
        item.crossDuplicate
      ) {
        type = "🔴🟠";
      } else if (item.localDuplicate) {
        type = "🔴";
      } else {
        type = "🟠";
      }

      return {
        Title: item.title,
        Documents: [...item.documents].join(", "),
        Count: item.count,
        Type: type,
      };
    });

    const workbook = XLSX.utils.book_new();

    const trackSheet =
      XLSX.utils.json_to_sheet(trackRows);

    const summarySheet =
      XLSX.utils.json_to_sheet(summaryRows);

    XLSX.utils.book_append_sheet(
      workbook,
      trackSheet,
      "Tracks"
    );

    XLSX.utils.book_append_sheet(
      workbook,
      summarySheet,
      "Duplicate Summary"
    );

    XLSX.writeFile(
      workbook,
      "track-title-report.xlsx"
    );
  }

  function matchesFilter(track) {
    const matchesSearch =
      !search ||
      track.title
        .toLowerCase()
        .includes(search.toLowerCase());

    if (!matchesSearch) return false;

    if (filter === "all") return true;

    if (filter === "local") {
      return track.duplicateInDocument;
    }

    if (filter === "cross") {
      return track.duplicateAcrossDocuments;
    }

    if (filter === "clean") {
      return (
        !track.duplicateInDocument &&
        !track.duplicateAcrossDocuments
      );
    }

    return true;
  }

  return (
    <div className="app">
      <header>
        <div>
          <h1>Track Title Collector</h1>
          <p>
            Paste documents → Extract titles → Check duplicates
          </p>
        </div>

        <button
          className="primary"
          onClick={() => setShowModal(true)}
        >
          + Paste Document
        </button>
      </header>

      <section className="stats">
        <div className="stat">
          <span>Documents</span>
          <strong>{documents.length}</strong>
        </div>

        <div className="stat">
          <span>Total tracks</span>
          <strong>{totalTracks}</strong>
        </div>

        <div className="stat">
          <span>Duplicate titles</span>
          <strong>{duplicateTitles.length}</strong>
        </div>
      </section>

      <section className="toolbar">
        <input
          type="text"
          placeholder="Search title..."
          value={search}
          onChange={(e) => setSearch(e.target.value)}
        />

        <select
          value={filter}
          onChange={(e) => setFilter(e.target.value)}
        >
          <option value="all">All</option>
          <option value="local">
            Duplicate in document
          </option>
          <option value="cross">
            Duplicate across documents
          </option>
          <option value="clean">No duplicates</option>
        </select>

        <button
          className="export"
          onClick={exportExcel}
        >
          Export Excel
        </button>
      </section>

      {documents.length === 0 ? (
        <div className="empty">
          <h2>No documents yet</h2>
          <p>
            Paste your first document to start extracting
            track titles.
          </p>

          <button
            className="primary"
            onClick={() => setShowModal(true)}
          >
            + Paste Document
          </button>
        </div>
      ) : (
        <div className="documents">
          {analyzedDocuments.map((doc) => {
            const visibleTracks = doc.tracks.filter(
              matchesFilter
            );

            return (
              <div className="document-card" key={doc.id}>
                <div className="document-header">
                  <div>
                    <h2>{doc.name}</h2>
                    <span>
                      {doc.tracks.length} tracks
                    </span>
                  </div>

                  <button
                    className="remove"
                    onClick={() =>
                      removeDocument(doc.id)
                    }
                  >
                    Remove
                  </button>
                </div>

                <div className="track-list">
                  {visibleTracks.length === 0 ? (
                    <div className="no-result">
                      No matching tracks.
                    </div>
                  ) : (
                    visibleTracks.map((track, index) => (
                      <div
                        className="track"
                        key={`${doc.id}-${index}`}
                      >
                        <span className="track-number">
                          {track.trackNumber}
                        </span>

                        <span className="track-title">
                          {track.title}
                        </span>

                        <span className="status">
                          {track.duplicateInDocument &&
                            "🔴"}

                          {track.duplicateAcrossDocuments &&
                            "🟠"}
                        </span>
                      </div>
                    ))
                  )}
                </div>
              </div>
            );
          })}
        </div>
      )}

      {duplicateTitles.length > 0 && (
        <section className="summary">
          <h2>Duplicate Summary</h2>

          {duplicateTitles.map((item) => (
            <div
              className="summary-row"
              key={normalizeTitle(item.title)}
            >
              <strong>{item.title}</strong>

              <span>
                {[...item.documents].join(", ")}
              </span>

              <span>{item.count} occurrences</span>

              <span>
                {item.localDuplicate && "🔴"}
                {item.crossDuplicate && "🟠"}
              </span>
            </div>
          ))}
        </section>
      )}

      {showModal && (
        <div
          className="modal-overlay"
          onClick={() => setShowModal(false)}
        >
          <div
            className="modal"
            onClick={(e) => e.stopPropagation()}
          >
            <div className="modal-header">
              <h2>Paste Document</h2>

              <button
                onClick={() => setShowModal(false)}
              >
                ×
              </button>
            </div>

            <label>Document name</label>

            <input
              value={documentName}
              onChange={(e) =>
                setDocumentName(e.target.value)
              }
              placeholder={`Document ${
                documents.length + 1
              }`}
            />

            <label>Document content</label>

            <textarea
              autoFocus
              value={documentText}
              onChange={(e) =>
                setDocumentText(e.target.value)
              }
              placeholder="Paste your document here..."
            />

            <div className="modal-actions">
              <button
                className="secondary"
                onClick={() =>
                  setShowModal(false)
                }
              >
                Cancel
              </button>

              <button
                className="primary"
                onClick={addDocument}
              >
                OK
              </button>
            </div>
          </div>
        </div>const regex =   /(?:^|\n)\s*(?:v\s*)?(?:[#>*`_\s]*)TRACK\s+(\d+)\s*:\s*(.+?)(?=\r?\n|$)/gi;
      )}
    </div>
  );
}

export default App;