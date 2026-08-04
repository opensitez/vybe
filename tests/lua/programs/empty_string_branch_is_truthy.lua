-- vybe-test: lua/programs/empty_string_branch_is_truthy
-- origin: languages/lua/tests/lua/test_programs.rs

local __w1 = "yes"
local __i = 0

if "" then do local __t = tostring("yes"); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end else do local __t = tostring("no"); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
