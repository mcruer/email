.onLoad <- function(libname, pkgname) {
  if (!"package:RDCOMClient" %in% search()) {
    attachNamespace("RDCOMClient")
  }
}
