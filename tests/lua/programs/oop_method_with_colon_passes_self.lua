-- vybe-test: lua/programs/oop_method_with_colon_passes_self
-- origin: languages/lua/tests/lua/test_programs.rs

local __w1 = "6"
local __i = 0

local obj = {v = 3}
function obj:double() return self.v * 2 end
do local __t = tostring(obj:double()); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
