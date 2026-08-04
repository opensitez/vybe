-- vybe-test: lua/metamethods/metamethod_index_chain_follows_metatable
-- origin: languages/lua/tests/lua/test_metamethods.rs

local __w1 = "1"
local __i = 0

local base={x=1}
local mid=setmetatable({}, {__index=base})
local top=setmetatable({}, {__index=mid})
do local __t = tostring(top.x); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
