const {onRequest} = require("firebase-functions/v2/https");
const OpenAI = require("openai");
require("dotenv").config();

exports.streamEndpoint = onRequest({ cors: true }, async (req, res) => {
  try {
    const { prompt } = req.body;
    const key = process.env.OPENAI_API_KEY;
    if (!key) {
      throw new Error("Missing OpenAI API Key");
    }
    const openai = new OpenAI({ apiKey: key });
    const completion = await openai.chat.completions.create({
      messages: [
        {role: "user", content: prompt},
      ],
      model: "gpt-4o-mini",
    });
    const completionContent = completion.choices[0].message.content;
    res.send(completionContent);
  } catch ( err ) {
    res.send(err)
  }
});
