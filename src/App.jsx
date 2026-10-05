import React, { useState, useEffect } from "react";
import * as XLSX from "xlsx-js-style";
import jsPDF from "jspdf";
import autoTable from "jspdf-autotable";

export default function App() {
  const [dados, setDados] = useState([]);
  useEffect(() => {
    fetch("/api/instrumentos")
      .then((res) => res.json())
      .then((data) => setDados(data))
      .catch(console.error);
  }, []);

  const [busca, setBusca] = useState("");
  const [editando, setEditando] = useState(null);

  const [ordem, setOrdem] = useState({
    campo: "",
    direcao: "asc",
  });

  const [paginaAtual, setPaginaAtual] = useState(1);
  const [itensPorPagina, setItensPorPagina] = useState(10);

  const [form, setForm] = useState({
    equipamento: "",
    fabricante: "",
    faixa: "",
    os: "",
    ativo: "",
    ci: "",
    certificado: "",
    calibracao: "",
    observacao: "",
  });

  const change = (e) => {
    setForm({
      ...form,
      [e.target.name]: e.target.value,
    });
  };

  const duplicar = (index) => {
    const copia = {
      ...dados[index],
      id: Date.now(),
    };

    setDados((prev) => [copia, ...prev]);

    alert("✅ Registro duplicado!");
  };

  const add = async () => {
    try {
      if (editando !== null) {
        setDados((prev) =>
          prev.map((item) =>
            item.id === editando ? { ...form, id: editando } : item,
          ),
        );

        setEditando(null);

        alert("✅ Registro atualizado!");
      } else {
        const res = await fetch("/api/instrumentos", {
          method: "POST",
          headers: {
            "Content-Type": "application/json",
          },
          body: JSON.stringify(form),
        });

        const novo = await res.json();

        setDados((prev) => [novo, ...prev]);

        alert("✅ Instrumento cadastrado!");
      }
    } catch (erro) {
      console.error(erro);
      alert("Erro ao salvar registro");
    }
  };

  const editar = (item) => {
    setForm(item);
    setEditando(item.id);
  };

  const remover = async (id) => {
    if (!window.confirm("Deseja excluir este registro?")) {
      return;
    }

    try {
      await fetch("/api/instrumentos", {
        method: "DELETE",
        headers: {
          "Content-Type": "application/json",
        },
        body: JSON.stringify({ id }),
      });

      setDados((prev) => prev.filter((item) => item.id !== id));

      alert("✅ Registro excluído!");
    } catch (erro) {
      console.error(erro);
      alert("Erro ao excluir");
    }
  };

  const limparTudo = () => {
    setDados([]);
  };

  const excel = () => {
    const dadosExcel = dados.map((item) => ({
      Equipamento: item.equipamento,
      Fabricante: item.fabricante,
      Faixa: item.faixa,
      OS: item.os,
      Ativo: item.ativo,
      CI: item.ci,
      Certificado: item.certificado,
      Calibração: item.calibracao || "-",
      Observação: item.observacao || "-",
    }));

    const ws = XLSX.utils.json_to_sheet(dadosExcel);

    ws["!cols"] = [
      { wch: 20 }, // Equipamento
      { wch: 20 }, // Fabricante
      { wch: 15 }, // Faixa
      { wch: 15 }, // OS
      { wch: 12 }, // Ativo
      { wch: 12 }, // CI
      { wch: 18 }, // Certificado
      { wch: 15 }, // Calibração
      { wch: 40 }, // Observação
    ];
    const range = XLSX.utils.decode_range(ws["!ref"]);

    for (let c = range.s.c; c <= range.e.c; c++) {
      const cellAddress = XLSX.utils.encode_cell({
        r: 0,
        c,
      });

      if (ws[cellAddress]) {
        ws[cellAddress].s = {
          fill: {
            fgColor: {
              rgb: "1F356D",
            },
          },
          font: {
            color: {
              rgb: "FFFFFF",
            },
            bold: true,
            sz: 12,
          },
          alignment: {
            horizontal: "center",
            vertical: "center",
            wrapText: true,
          },
        };
      }
    }

    const wb = XLSX.utils.book_new();

    XLSX.utils.book_append_sheet(wb, ws, "Instrumentos");
    const hoje = new Date().toLocaleDateString("pt-BR").replaceAll("/", "-");

    XLSX.writeFile(wb, `Cadastro_Instrumentos_${hoje}.xlsx`);
  };

  const gerarPDF = () => {
    const doc = new jsPDF("landscape");

    // Título
    doc.setFontSize(22);
    doc.setTextColor(31, 53, 109);
    doc.text("Cadastro de Instrumentos", 14, 20);

    doc.setDrawColor(31, 53, 109);
    doc.setLineWidth(0.6);
    doc.line(14, 23, 280, 23);

    // Data de emissão
    doc.setFontSize(10);
    doc.setTextColor(100);
    doc.text(`Emitido em: ${new Date().toLocaleDateString("pt-BR")}`, 14, 28);
    doc.setFontSize(11);
    doc.setTextColor(80);

    doc.text(`Total de Registros: ${dados.length}`, 14, 34);
    doc.setDrawColor(31, 53, 109);
    doc.setLineWidth(0.8);
    doc.line(14, 38, 280, 38);

    autoTable(doc, {
      startY: 46,

      head: [
        [
          "Equipamento",
          "Fabricante",
          "Faixa",
          "OS",
          "Ativo",
          "CI",
          "Certificado",
          "Calibração",
          "Observação",
        ],
      ],

      body: dados.map((item) => [
        item.equipamento,
        item.fabricante,
        item.faixa,
        item.os,
        item.ativo,
        item.ci,
        item.certificado,
        item.calibracao || "-",
        item.observacao || "-",
      ]),

      headStyles: {
        fillColor: [31, 53, 109],
        textColor: [255, 255, 255],
        fontStyle: "bold",
        halign: "center",
      },

      alternateRowStyles: {
        fillColor: [245, 247, 250],
      },

      styles: {
        fontSize: 9,
        cellPadding: 4,
        overflow: "linebreak",
        valign: "middle",
      },

      columnStyles: {
        0: { cellWidth: 25 },
        1: { cellWidth: 25 },
        2: { cellWidth: 20 },
        8: { cellWidth: 70 },
      },

      margin: {
        top: 35,
      },
    });

    // Rodapé
    const paginas = doc.internal.getNumberOfPages();

    for (let i = 1; i <= paginas; i++) {
      doc.setPage(i);

      doc.setFontSize(9);
      doc.setTextColor(120);

      doc.text(
        "Sistema de Cadastro de Instrumentos",
        14,
        doc.internal.pageSize.height - 10,
      );

      doc.text(
        `Página ${i} de ${paginas}`,
        doc.internal.pageSize.width - 40,
        doc.internal.pageSize.height - 10,
      );
    }

    const hoje = new Date().toLocaleDateString("pt-BR").replaceAll("/", "-");

    doc.save(`Cadastro_Instrumentos_${hoje}.pdf`);
  };
  const ordenar = (campo) => {
    setOrdem((prev) => ({
      campo,
      direcao: prev.campo === campo && prev.direcao === "asc" ? "desc" : "asc",
    }));
  };

  const dadosFiltrados = dados
    .filter((item) =>
      Object.values(item).some((valor) =>
        String(valor).toLowerCase().includes(busca.toLowerCase()),
      ),
    )
    .sort((a, b) => {
      if (!ordem.campo) return 0;

      const valorA = String(a[ordem.campo] || "").toLowerCase();
      const valorB = String(b[ordem.campo] || "").toLowerCase();

      if (ordem.direcao === "asc") {
        return valorA.localeCompare(valorB);
      }

      return valorB.localeCompare(valorA);
    });

  const indiceFinal = paginaAtual * itensPorPagina;
  const indiceInicial = indiceFinal - itensPorPagina;

  const dadosPaginados = dadosFiltrados.slice(indiceInicial, indiceFinal);

  const totalPaginas = Math.ceil(dadosFiltrados.length / itensPorPagina);

  const inputStyle = {
    width: "100%",
    padding: "14px",
    boxSizing: "border-box",
    borderRadius: "12px",
    border: "1px solid #dbe2ea",
    background: "#fafbfd",
    outline: "none",
    fontSize: "15px",
  };

  return (
    <div
      style={{
        background: "#f3f4f6",
        minHeight: "100vh",
        padding: "20px",
        display: "flex",
        justifyContent: "center",
        alignItems: "flex-start",
      }}
    >
      <div
        style={{
          maxWidth: "1000px",
          width: "100%",
          padding: "15px",
          fontFamily: "Arial, sans-serif",
          boxSizing: "border-box",
          background: "#fff",
          borderRadius: "28px",
          boxShadow: "0 12px 30px rgba(31,53,109,0.12)",
          border: "1px solid #eef2f7",
        }}
      >
        <h1
          style={{
            textAlign: "center",
            fontSize: "30px",
            color: "#1F356D",
            fontWeight: "700",
            lineHeight: "1.2",
          }}
        >
          📋 Cadastro de Instrumentos
        </h1>

        <div
          style={{
            display: "grid",
            gridTemplateColumns: "1fr",
            gap: "12px",
            marginTop: "30px",
          }}
        >
          <input
            name="equipamento"
            placeholder="Equipamento"
            value={form.equipamento}
            onChange={change}
            style={inputStyle}
          />

          <input
            name="fabricante"
            placeholder="Fabricante"
            value={form.fabricante}
            onChange={change}
            style={inputStyle}
          />

          <input
            name="faixa"
            placeholder="Faixa"
            value={form.faixa}
            onChange={change}
            style={inputStyle}
          />

          <input
            name="os"
            placeholder="OS"
            value={form.os}
            onChange={change}
            style={inputStyle}
          />

          <input
            name="ativo"
            placeholder="Ativo"
            value={form.ativo}
            onChange={change}
            style={inputStyle}
          />

          <input
            name="ci"
            placeholder="CI"
            value={form.ci}
            onChange={change}
            style={inputStyle}
          />

          <input
            name="certificado"
            placeholder="Certificado"
            value={form.certificado}
            onChange={change}
            style={inputStyle}
          />

          <input
            type="date"
            name="calibracao"
            value={form.calibracao}
            onChange={change}
            style={inputStyle}
          />
          <textarea
            name="observacao"
            placeholder="Observação"
            value={form.observacao}
            onChange={change}
            rows="4"
            style={{
              ...inputStyle,
              resize: "vertical",
            }}
          />
        </div>

        <div
          style={{
            display: "flex",
            flexDirection: "column",
            gap: "10px",
            marginTop: "20px",
          }}
        >
          <button
            onClick={add}
            style={{
              background: "#1F356D",
              color: "white",
              border: "none",
              padding: "14px",
              fontSize: "16px",
              fontWeight: "bold",
              borderRadius: "5px",
              cursor: "pointer",
              width: "100%",
              letterSpacing: "0.3px",
            }}
          >
            {editando !== null ? "Salvar Alteração" : "Adicionar"}
          </button>

          <button
            onClick={excel}
            style={{
              background: "#6CC2E5",
              color: "#1F356D",
              fontWeight: "bold",
              border: "none",
              padding: "10px 20px",
              borderRadius: "5px",
              cursor: "pointer",
              width: "100%",
            }}
          >
            Gerar Excel
          </button>
          <button
            onClick={gerarPDF}
            style={{
              background: "#1F356D",
              color: "white",
              border: "none",
              padding: "14px",
              borderRadius: "10px",
              cursor: "pointer",
              width: "100%",
              fontWeight: "bold",
            }}
          >
            Gerar PDF
          </button>
          <button
            onClick={limparTudo}
            style={{
              background: "#d64545",
              color: "white",
              border: "none",
              padding: "10px 20px",
              borderRadius: "5px",
              cursor: "pointer",
              width: "100%",
            }}
          >
            Limpar Tudo
          </button>
        </div>

        <h2
          style={{
            marginTop: "25px",
            color: "#1F356D",
          }}
        >
          📋 Total de Registros: {dados.length}
        </h2>

        <p
          style={{
            color: "#666",
            marginTop: "-5px",
          }}
        >
          🔎 Resultados encontrados: {dadosFiltrados.length}
        </p>
        <p
          style={{
            color: "#666",
            marginTop: "-5px",
          }}
        >
          📄 Exibindo {itensPorPagina} registros por página
        </p>
        <div style={{ marginBottom: "15px" }}>
          <label
            style={{
              marginRight: "10px",
              fontWeight: "bold",
              color: "#1F356D",
            }}
          >
            Registros por página:
          </label>

          <select
            value={itensPorPagina}
            onChange={(e) => {
              setItensPorPagina(Number(e.target.value));
              setPaginaAtual(1);
            }}
            style={{
              padding: "8px",
              borderRadius: "6px",
              border: "1px solid #ccc",
            }}
          >
            <option value={5}>5</option>
            <option value={10}>10</option>
            <option value={25}>25</option>
            <option value={50}>50</option>
          </select>
        </div>
        <input
          type="text"
          placeholder="Pesquisar equipamento..."
          value={busca}
          onChange={(e) => {
            setBusca(e.target.value);
            setPaginaAtual(1);
          }}
          style={{
            width: "100%",
            padding: "10px",
            marginBottom: "20px",
            boxSizing: "border-box",
          }}
        />

        <div style={{ overflowX: "auto", marginTop: "15px" }}>
          <table
            style={{
              width: "100%",
              minWidth: "800px",
              borderCollapse: "collapse",
              backgroundColor: "#fff",
              boxShadow: "0 2px 8px rgba(0,0,0,0.1)",
            }}
          >
            <thead>
              <tr
                style={{
                  backgroundColor: "#1F356D",
                  color: "#fff",
                }}
              >
                <th
                  onClick={() => ordenar("equipamento")}
                  style={{
                    padding: "12px",
                    cursor: "pointer",
                  }}
                >
                  Equipamento
                </th>{" "}
                <th
                  onClick={() => ordenar("fabricante")}
                  style={{
                    padding: "12px",
                    cursor: "pointer",
                  }}
                >
                  Fabricante
                </th>
                <th
                  onClick={() => ordenar("faixa")}
                  style={{
                    padding: "12px",
                    cursor: "pointer",
                  }}
                >
                  Faixa
                </th>
                <th
                  onClick={() => ordenar("os")}
                  style={{
                    padding: "12px",
                    cursor: "pointer",
                  }}
                >
                  OS
                </th>
                <th
                  onClick={() => ordenar("ativo")}
                  style={{
                    padding: "12px",
                    cursor: "pointer",
                  }}
                >
                  Ativo
                </th>
                <th
                  onClick={() => ordenar("ci")}
                  style={{
                    padding: "12px",
                    cursor: "pointer",
                  }}
                >
                  CI
                </th>
                <th
                  onClick={() => ordenar("certificado")}
                  style={{
                    padding: "12px",
                    cursor: "pointer",
                  }}
                >
                  Certificado
                </th>
                <th style={{ padding: "12px" }}>Calibração</th>
                <th style={{ padding: "12px" }}>Observação</th>
                <th style={{ padding: "12px" }}>Ações</th>
              </tr>
            </thead>

            <tbody>
              {dadosFiltrados.length === 0 ? (
                <tr>
                  <td
                    colSpan="10"
                    style={{
                      padding: "20px",
                      textAlign: "center",
                    }}
                  >
                    Nenhum registro encontrado.
                  </td>
                </tr>
              ) : (
                dadosPaginados.map((d, i) => (
                  <tr
                    key={i}
                    style={{
                      backgroundColor: i % 2 === 0 ? "#ffffff" : "#f8fafc",
                    }}
                  >
                    <td style={{ padding: "10px", textAlign: "center" }}>
                      {d.equipamento}
                    </td>

                    <td style={{ padding: "10px", textAlign: "center" }}>
                      {d.fabricante}
                    </td>

                    <td style={{ padding: "10px", textAlign: "center" }}>
                      {d.faixa}
                    </td>

                    <td style={{ padding: "10px", textAlign: "center" }}>
                      {d.os}
                    </td>

                    <td style={{ padding: "10px", textAlign: "center" }}>
                      {d.ativo}
                    </td>

                    <td style={{ padding: "10px", textAlign: "center" }}>
                      {d.ci}
                    </td>

                    <td style={{ padding: "10px", textAlign: "center" }}>
                      {d.certificado}
                    </td>

                    <td style={{ padding: "10px", textAlign: "center" }}>
                      {d.calibracao || "-"}
                    </td>

                    <td
                      style={{
                        padding: "10px",
                        minWidth: "220px",
                        whiteSpace: "normal",
                        wordBreak: "break-word",
                      }}
                    >
                      {d.observacao || "-"}
                    </td>

                    <td
                      style={{
                        padding: "10px",
                        textAlign: "center",
                        whiteSpace: "nowrap",
                      }}
                    >
                      <button
                        onClick={() => duplicar(i)}
                        style={{
                          background: "#2563eb",
                          color: "white",
                          border: "none",
                          padding: "6px 12px",
                          borderRadius: "5px",
                          cursor: "pointer",
                          marginRight: "5px",
                        }}
                      >
                        Duplicar
                      </button>

                      <button
                        onClick={() => editar(d)}
                        style={{
                          background: "#f59e0b",
                          color: "white",
                          border: "none",
                          padding: "6px 12px",
                          borderRadius: "5px",
                          cursor: "pointer",
                          marginRight: "5px",
                        }}
                      >
                        Editar
                      </button>

                      <button
                        onClick={() => remover(d.id)}
                        style={{
                          background: "#dc2626",
                          color: "white",
                          border: "none",
                          padding: "6px 12px",
                          borderRadius: "5px",
                          cursor: "pointer",
                        }}
                      >
                        Excluir
                      </button>
                    </td>
                  </tr>
                ))
              )}
            </tbody>
          </table>
          <div
            style={{
              display: "flex",
              justifyContent: "center",
              alignItems: "center",
              gap: "15px",
              marginTop: "20px",
              flexWrap: "wrap",
            }}
          >
            <button
              disabled={paginaAtual === 1}
              onClick={() => setPaginaAtual(paginaAtual - 1)}
              style={{
                padding: "8px 15px",
                borderRadius: "6px",
                border: "none",
                background: "#1F356D",
                color: "white",
                cursor: paginaAtual === 1 ? "not-allowed" : "pointer",
                fontWeight: "bold",
                opacity: paginaAtual === 1 ? 0.5 : 1,
              }}
            >
              ◀ Anterior
            </button>

            <div
              style={{
                display: "flex",
                gap: "5px",
                flexWrap: "wrap",
              }}
            >
              {Array.from({ length: totalPaginas }, (_, i) => i + 1).map(
                (pagina) => (
                  <button
                    key={pagina}
                    onClick={() => setPaginaAtual(pagina)}
                    style={{
                      padding: "6px 12px",
                      border: "none",
                      borderRadius: "5px",
                      cursor: "pointer",
                      background:
                        paginaAtual === pagina ? "#1F356D" : "#e5e7eb",
                      color: paginaAtual === pagina ? "white" : "#111827",
                      fontWeight: "bold",
                      transition: "0.2s",
                    }}
                  >
                    {pagina}
                  </button>
                ),
              )}
            </div>

            <button
              disabled={paginaAtual === totalPaginas || totalPaginas === 0}
              onClick={() => setPaginaAtual(paginaAtual + 1)}
              style={{
                padding: "8px 15px",
                borderRadius: "6px",
                border: "none",
                background: "#1F356D",
                color: "white",
                cursor:
                  paginaAtual === totalPaginas ? "not-allowed" : "pointer",
                fontWeight: "bold",
                opacity: paginaAtual === totalPaginas ? 0.5 : 1,
              }}
            >
              Próxima ▶
            </button>
          </div>
        </div>
      </div>
    </div>
  );
}
