# vybe-test: python/python_selectors_io_mux/test_selectors_kqueue_selector_available_on_bsd
# origin: languages/python/tests/python/test_python_selectors_io_mux.rs

import selectors, sys
if sys.platform in ("darwin", "freebsd"):
    print(hasattr(selectors, "KqueueSelector"))
else:
    print(True)
