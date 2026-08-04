-- vybe-test: lua/environment/function_inherits_defining_environment
-- origin: languages/lua/tests/lua/test_environment.rs

local _ENV = {print = print, v = 1}
local function f() return v end
print(f())
