-- vybe-test: lua/lexical_environments_advanced/test_env_set_upvalue_via_debug
-- origin: languages/lua/tests/lua/test_lexical_environments_advanced.rs

local __w1 = "42"
local __i = 0

local function f() return x end
local name, val = debug.getupvalue(f, 1)
if name == '_ENV' then
    debug.setupvalue(f, 1, {x = 42})
end
do local __t = tostring(f()); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
