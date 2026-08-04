-- vybe-test: lua/loops_for_generic/generic_for_sorted_keys_iteration
-- origin: languages/lua/tests/lua/test_loops_for_generic.rs

local __w1 = "a1b2c3"
local __i = 0

local t = {c=3, a=1, b=2}
local keys = {}
for k in pairs(t) do keys[#keys+1] = k end
table.sort(keys)
local s = ''
for _, k in ipairs(keys) do s = s .. k .. t[k] end
do local __t = tostring(s); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
