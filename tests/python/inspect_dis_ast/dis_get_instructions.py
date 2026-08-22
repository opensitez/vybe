# vybe-test: python/inspect_dis_ast/dis_get_instructions
# origin: languages/python/tests/python/test_inspect_dis_ast.rs

import dis
c = compile('1', '<s>', 'eval')
list(dis.get_instructions(c))
