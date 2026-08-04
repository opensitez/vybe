-- vybe-test: lua/language_semantics_extended/control_if_elseif
-- origin: languages/lua/tests/lua/test_language_semantics_extended.rs

local __w1 = "elseif"
local __i = 0

local x
if false then x = "if" elseif true then x = "elseif" else x = "else" end
do local __t = tostring(x); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
