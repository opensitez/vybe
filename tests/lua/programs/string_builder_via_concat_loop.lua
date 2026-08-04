-- vybe-test: lua/programs/string_builder_via_concat_loop
-- origin: languages/lua/tests/lua/test_programs.rs

local __w1 = "abc"
local __i = 0

local parts = {"a", "b", "c"}
local s = ""
for i = 1, #parts do s = s .. parts[i] end
do local __t = tostring(s); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
