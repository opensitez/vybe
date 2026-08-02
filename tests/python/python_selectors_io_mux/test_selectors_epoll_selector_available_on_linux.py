# vybe-test: python/python_selectors_io_mux/test_selectors_epoll_selector_available_on_linux
# origin: languages/python/tests/python/test_python_selectors_io_mux.rs

import selectors, sys
if sys.platform == "linux":
    print(hasattr(selectors, "EpollSelector"))
else:
    print(True)
