-- vybe-test: lua/programs/ends_with_suffix_check
-- origin: languages/lua/tests/lua/test_programs.rs

local __w1 = "true"
local __i = 0

local s = "file.lua"
do local __t = tostring(string.sub(s, -4) == ".lua"); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
