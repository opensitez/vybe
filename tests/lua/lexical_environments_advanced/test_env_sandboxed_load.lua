-- vybe-test: lua/lexical_environments_advanced/test_env_sandboxed_load
-- origin: languages/lua/tests/lua/test_lexical_environments_advanced.rs

local __w1 = "nil"
local __i = 0

local f = load('return math and type(math) or "nil"', '', 't', {})
do local __t = tostring(f()); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
