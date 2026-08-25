# Current
- --engine to allow multiple engines, right now html bo
- note: for now, you need to manually place htmlbox in crates
  - cd <vybe>/crates;git clone https://github.com/opensitez/HtmlBox
  - when htmlbox is finalized, it will be pulled through crates.io
* Finalize HtmlBox interface and public to crates.io
* Finalize Vybe Widgets, Extract and publish to github
* WASM3 Compliance


# 20260825/ v0.6.1
- improved whatwg compatibility
- made the browser seam compatible both with htmlbox and vybe_widgets
- vybe_widget made compatible
- cm3 component model
- wasi 0.3.1
- removed wasi:io
- experimental support for htmlbox
- namespace case sensitivity is now fixed at compile time
- compiler tests validated for extractor dammage
- test extractor deleted as tests fixed
- old tests deleted
- no more static vars in compiler

# 2026/06/ v0.6.0
- vybe_host is now split into platforms/web platforms/ecma platforms/node
- languages are now split into their own crate in languages/<lang>
- tests are now actual files instead of rustcode to avoid recompilation/run in real lang
- testrunner allow to run new tests/multi workers/no hang
- testrunner allow to run in other compilers/check validity
- gui are now html, gui adapted through primitives/gui.rs
- common tree resolver/ use_dotnet is dead
- vybe_emitter is not vybe_compiler/src/primitives
- common ast widened to allow more concepts, directives
- vybe:gui is deleted, replaced by web:* whatwg interface
- vybe now has a dom and behaves like a dom
- lua compiler

# 202603/ v0.5.0
- wasi sql replaced vybe:sql
- --serve mode to allow serving php
- rewritten compilers as pest grammar
- cobol/fortran compiler
