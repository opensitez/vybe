-- vybe-test: lua/goto/goto_over_global_function_definition_is_valid
-- origin: languages/lua/tests/lua/test_goto.rs

local __w1 = "function"
local __i = 0

goto after
function skip_fn() return 99 end
::after::
do local __t = tostring(type(skip_fn)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
