require("dotenv").config();
const express = require("express");
const session = require("express-session");
const authRoutes = require("./routes/auth.routes");
const fileRoutes = require("./routes/file.routes");
const versionRoutes = require("./routes/version.routes");

const app = express();
app.use(express.json());
app.use(express.urlencoded({ extended: true }));
app.use(session({ secret: process.env.SESSION_SECRET, resave: false, saveUninitialized: false }));

app.use("/", authRoutes);
app.use("/", fileRoutes);
app.use("/", versionRoutes);

app.listen(process.env.PORT || 3000, () => {
  console.log(`Server running at http://localhost:${process.env.PORT || 3000}`);
});