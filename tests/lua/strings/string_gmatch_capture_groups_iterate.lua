-- vybe-test: lua/strings/string_gmatch_capture_groups_iterate
-- origin: languages/lua/tests/lua/test_strings.rs

local __w1 = "1+2"
local __i = 0

local s = ""
for a,b in string.gmatch("1=2", "(%d+)=(%d+)") do s = a .. "+" .. b end
do local __t = tostring(s); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
