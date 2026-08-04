-- vybe-test: lua/environments_ext/test_env_upvalue_mutation
-- origin: languages/lua/tests/lua/test_environments_ext.rs

local _ENV = {print=print, a=10}; local function f() a=20 end; f(); print(a)
