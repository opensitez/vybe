-- vybe-test: lua/env_lexical_binding/env_lexical_resolve
-- origin: languages/lua/tests/lua/test_env_lexical_binding.rs

local _ENV = {print=print, x=100}
print(x)
