-- vybe-test: lua/goto_advanced/goto_decl_label
-- origin: languages/lua/tests/lua/test_goto_advanced.rs

local __w1 = "1"
local __i = 0

local x = 1
goto lbl
x = 2
::lbl::
do local __t = tostring(x); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
