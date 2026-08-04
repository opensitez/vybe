-- vybe-test: lua/lexical_environments_advanced/test_env_change_upvalue
-- origin: languages/lua/tests/lua/test_lexical_environments_advanced.rs

local _ENV = {print=print, type=type, debug=debug, load=load}
local function get_a() return a end
_ENV.a = 100
print(get_a())
