-- vybe-test: lua/select_builtin/select_count_nils_only
-- origin: languages/lua/tests/lua/test_select_builtin.rs

local __w1 = "2"
local __i = 0

local function f(...) return select('#', ...) end
do local __t = tostring(f(nil, nil)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
