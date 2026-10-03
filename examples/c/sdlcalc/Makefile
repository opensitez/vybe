# Multi-file SDL calculator. Exercises: gated includes (SDL2/SDL.h, math.h,
# stdio.h), mouse+keyboard events, multiple translation units, entry
# override, wasm artifact.
#
#   make run   - interactive calculator (window)
#   make test  - headless input self-test via SDL_PushEvent (--entry)
#   make wasm  - build calc.wasm artifact
#   make clean

VYBEX ?= ../../../target/debug/vybex
VYBEXREL ?= ../../../target/release/vybex
SRCS  := main.c calc.c ui.c
TSRCS := selftest.c calc.c ui.c
HDRS  := calc.h ui.h

.PHONY: run test demo shot clean

run: $(SRCS) $(HDRS)
	$(VYBEX) $(SRCS)

demorel: $(SRCS) $(HDRS)
	$(VYBEXREL) $(SRCS)

test: $(TSRCS) $(HDRS)
	$(VYBEX) --entry self_test $(TSRCS)

calc.wasm: $(SRCS) $(HDRS)
	$(VYBEX) -w $(SRCS)
	mv main.wasm calc.wasm

wasm: calc.wasm

# One-frame demo: window shows the rendered calculator (display = 3.75)
# after the program returns.
demo: demo.c calc.c ui.c $(HDRS)
	$(VYBEX) demo.c calc.c ui.c

# Same frame, written offscreen to calc.png (no window).
shot: demo.c calc.c ui.c $(HDRS)
	$(VYBEX) --capture calc.png demo.c calc.c ui.c

clean:
	rm -f calc.wasm main.wasm calc.png
