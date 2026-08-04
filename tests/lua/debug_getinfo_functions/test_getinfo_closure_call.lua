-- vybe-test: lua/debug_getinfo_functions/test_getinfo_closure_call
-- origin: languages/lua/tests/lua/test_debug_getinfo_functions.rs

local __w1 = "3"
local __i = 0

local function mk(v)
  return function() return v end
end
local c = mk(3)
do local __t = tostring(c()); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
