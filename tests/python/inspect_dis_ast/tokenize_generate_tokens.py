# vybe-test: python/inspect_dis_ast/tokenize_generate_tokens
# origin: languages/python/tests/python/test_inspect_dis_ast.rs

import tokenize
import io
tokenize.generate_tokens(io.StringIO('x=1').readline)
