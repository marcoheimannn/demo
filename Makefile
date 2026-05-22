LATEX ?= pdflatex
LATEXFLAGS ?= -interaction=nonstopmode -halt-on-error

.PHONY: all clean

all: diagram.pdf diagram_auditeo.pdf

diagram.pdf: diagram.tex
	$(LATEX) $(LATEXFLAGS) diagram.tex

clean:
	rm -f diagram.aux diagram.log diagram.pdf diagram_auditeo.aux diagram_auditeo.log diagram_auditeo.pdf


diagram_auditeo.pdf: diagram_auditeo.tex
	$(LATEX) $(LATEXFLAGS) diagram_auditeo.tex
