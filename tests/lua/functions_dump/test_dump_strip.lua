-- vybe-test: lua/functions_dump/test_dump_strip
-- origin: languages/lua/tests/lua/test_functions_dump.rs

local __w1 = "true"
local __i = 0

local f = function() return 42 end; local d1 = string.dump(f); local d2 = string.dump(f, true); do local __t = tostring(tostring(string.len(d2) <= string.len(d1))); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
