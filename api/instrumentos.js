import { neon } from "@neondatabase/serverless";

const sql = neon(process.env.DATABASE_URL);

export default async function handler(req, res) {
  try {
    if (req.method === "GET") {
      const instrumentos = await sql`
        SELECT *
        FROM instrumentos
        ORDER BY id DESC
      `;

      return res.status(200).json(instrumentos);
    }

    res.status(405).json({
      error: "Método não permitido",
    });
  } catch (error) {
    res.status(500).json({
      error: error.message,
    });
  }
}