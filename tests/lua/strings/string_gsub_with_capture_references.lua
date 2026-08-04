-- vybe-test: lua/strings/string_gsub_with_capture_references
-- origin: languages/lua/tests/lua/test_strings.rs

local __w1 = "20/10\t1"
local __i = 0

do local __t = tostring(string.gsub("10-20", "(%d+)-(%d+)", "%2/%1")); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
