import { neon } from "@neondatabase/serverless";

const sql = neon(process.env.DATABASE_URL);

export default async function handler(req, res) {
  try {
    const resultado = await sql`
      SELECT COUNT(*) as total
      FROM instrumentos
    `;

    return res.status(200).json(resultado);
  } catch (error) {
    return res.status(500).json({
      erro: error.message,
    });
  }
}