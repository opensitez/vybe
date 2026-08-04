-- vybe-test: lua/iteration/pairs_reads_key_value_from_hash
-- origin: languages/lua/tests/lua/test_iteration.rs

local __w1 = "name=lua"
local __i = 0

local t = {name = "lua"}
for k, v in pairs(t) do do local __t = tostring(k .. "=" .. v); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
