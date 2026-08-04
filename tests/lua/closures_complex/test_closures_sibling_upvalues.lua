-- vybe-test: lua/closures_complex/test_closures_sibling_upvalues
-- origin: languages/lua/tests/lua/test_closures_complex.rs

local __w1 = "1 2 1"
local __i = 0

local function make_counter()
    local count = 0
    return function() count = count + 1 return count end,
           function() count = count - 1 return count end
end
local inc, dec = make_counter()
do local __t = tostring(inc() .. ' ' .. inc() .. ' ' .. dec()); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
