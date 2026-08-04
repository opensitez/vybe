-- vybe-test: lua/functions_multiple_returns/test_multret_unpack_middle
-- origin: languages/lua/tests/lua/test_functions_multiple_returns.rs

local __w1 = "2"
local __i = 0

local function g(...) return select('#', ...) end; do local __t = tostring(g(table.unpack({1,2}), 3)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
