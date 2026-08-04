-- vybe-test: lua/iteration/pairs_visits_hash_keys
-- origin: languages/lua/tests/lua/test_iteration.rs

local __w1 = "2"
local __i = 0

local t={x=1,y=2}
local n=0
for _ in pairs(t) do n=n+1 end
do local __t = tostring(n); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
