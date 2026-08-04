-- vybe-test: lua/closures/upvalue_not_visible_to_sibling_inner_functions
-- origin: languages/lua/tests/lua/test_closures.rs

local __w1 = "1"
local __i = 0

local n=0
local function a() n=1 end
local function b() return n end
a()
do local __t = tostring(b()); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
