function ensureAuthenticated(req, res, next) {
	console.log("selectedFlag:", req.session);
  if (!req.session.user) {
    return res.redirect("/login?flag="+req.session.selectedFlag);
  }
  next();
}

module.exports = { ensureAuthenticated };
