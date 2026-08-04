-- vybe-test: lua/env_lexical_binding/env_fn_inherit
-- origin: languages/lua/tests/lua/test_env_lexical_binding.rs

local _ENV = {print=print, y=42}
local function f() return y end
print(f())
