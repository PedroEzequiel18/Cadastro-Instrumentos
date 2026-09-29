import { neon } from "@neondatabase/serverless";

const sql = neon(process.env.DATABASE_URL);

export default async function handler(req, res) {
  try {
    // LISTAR
    if (req.method === "GET") {
      const instrumentos = await sql`
        SELECT *
        FROM instrumentos
        ORDER BY id DESC
      `;

      return res.status(200).json(instrumentos);
    }

    // CADASTRAR
    if (req.method === "POST") {
      const {
        equipamento,
        fabricante,
        faixa,
        os,
        ativo,
        ci,
        certificado,
        calibracao,
        observacao,
      } = req.body;

      const resultado = await sql`
        INSERT INTO instrumentos (
          equipamento,
          fabricante,
          faixa,
          os,
          ativo,
          ci,
          certificado,
          calibracao,
          observacao
        )
        VALUES (
          ${equipamento},
          ${fabricante},
          ${faixa},
          ${os},
          ${ativo},
          ${ci},
          ${certificado},
          ${calibracao || null},
          ${observacao}
        )
        RETURNING *
      `;

      return res.status(201).json(resultado[0]);
    }

    // EXCLUIR
    if (req.method === "DELETE") {
      const { id } = req.body;

      await sql`
        DELETE FROM instrumentos
        WHERE id = ${Number(id)}
      `;

      return res.status(200).json({
        success: true,
        deletedId: id,
      });
    }

    return res.status(405).json({
      error: "Método não permitido",
    });
  } catch (error) {
    return res.status(500).json({
      error: error.message,
    });
  }
}