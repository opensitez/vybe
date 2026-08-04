-- vybe-test: lua/functions/colon_method_syntax_passes_self_as_first_arg
-- origin: languages/lua/tests/lua/test_functions.rs

local __w1 = "10"
local __i = 0

local obj = {value = 10}
function obj:get() return self.value end
do local __t = tostring(obj:get()); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
