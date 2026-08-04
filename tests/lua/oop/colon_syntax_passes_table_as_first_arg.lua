-- vybe-test: lua/oop/colon_syntax_passes_table_as_first_arg
-- origin: languages/lua/tests/lua/test_oop.rs

local __w1 = "2"
local __i = 0

local t = {}
function t.add(self, x) self.v = (self.v or 0) + x end
t:add(2)
do local __t = tostring(t.v); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
