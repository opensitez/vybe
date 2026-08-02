# vybe-test: python/python_symtable_symbol_tables/test_symtable_get_symbols_in_scope
# origin: languages/python/tests/python/test_python_symtable_symbol_tables.rs

import symtable
st = symtable.symtable("x = 10; y = 20", "<string>", "exec")
sym_names = [s.get_name() for s in st.get_symbols()]
print("x" in sym_names and "y" in sym_names)
