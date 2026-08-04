-- vybe-test: lua/env_lexical_binding/env_do_block_scope
-- origin: languages/lua/tests/lua/test_env_lexical_binding.rs

x = 5
do
  local _ENV = {print=print, x=10}
  print(x)
end
print(x)
