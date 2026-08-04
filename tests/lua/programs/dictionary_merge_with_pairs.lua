-- vybe-test: lua/programs/dictionary_merge_with_pairs
-- origin: languages/lua/tests/lua/test_programs.rs

local __w1 = "3"
local __i = 0

local a, b, out = {x = 1}, {y = 2}, {}
for k, v in pairs(a) do out[k] = v end
for k, v in pairs(b) do out[k] = v end
do local __t = tostring(out.x + out.y); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
