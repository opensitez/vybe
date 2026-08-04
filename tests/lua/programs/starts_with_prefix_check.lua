-- vybe-test: lua/programs/starts_with_prefix_check
-- origin: languages/lua/tests/lua/test_programs.rs

local __w1 = "true"
local __i = 0

local s = "lua.lang"
do local __t = tostring(string.sub(s, 1, 3) == "lua"); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
