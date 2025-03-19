export default function handler(req, res) {
    if (req.method === "POST") {
        console.log("Recebendo evento do HubSpot:", req.body);
        res.status(200).json({ message: "Webhook recebido com sucesso!" });
    } else {
        res.status(405).json({ message: "Método não permitido" });
    }
}

