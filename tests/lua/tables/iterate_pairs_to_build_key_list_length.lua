-- vybe-test: lua/tables/iterate_pairs_to_build_key_list_length
-- origin: languages/lua/tests/lua/test_tables.rs

local __w1 = "3"
local __i = 0

local t = {x=1, y=2, z=3}
local n = 0
for _ in pairs(t) do n = n + 1 end
do local __t = tostring(n); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
