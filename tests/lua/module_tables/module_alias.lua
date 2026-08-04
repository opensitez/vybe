-- vybe-test: lua/module_tables/module_alias
-- origin: languages/lua/tests/lua/test_module_tables.rs

local __w1 = "49"
local __i = 0

local MyMath = {}
function MyMath.sq(n) return n * n end
local sq = MyMath.sq
do local __t = tostring(sq(7)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
