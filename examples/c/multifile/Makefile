# Standard multi-file C project driven by make, compiled by vybex.
#
# What works today:
#   make run       - compile all TUs and run (linker finds `main`)
#   make main.wasm - build the linked .wasm artifact; only rebuilds when a
#                    source or header changed (ordinary make timestamps)
#   make run-wasm  - execute the built artifact
#   make clean
#
# Entry override, ld -e style:  make run ENTRY=some_function
#
# Not yet: per-TU objects (vybex -c foo.c -o foo.wasmo) — the cc-style
# incremental object flow lands with the .wasmo linker step. Until then the
# "link" recompiles all sources, which make still gates on timestamps.

VYBEX ?= ../../../target/debug/vybex
SRCS  := main.c geometry.c counter.c
HDRS  := geometry.h counter.h
ENTRY ?=
ENTRYFLAG := $(if $(ENTRY),--entry $(ENTRY),)

.PHONY: run run-wasm clean

run: $(SRCS) $(HDRS)
	$(VYBEX) $(ENTRYFLAG) $(SRCS)

main.wasm: $(SRCS) $(HDRS)
	$(VYBEX) -w $(SRCS)

run-wasm: main.wasm
	$(VYBEX) main.wasm

clean:
	rm -f main.wasm
