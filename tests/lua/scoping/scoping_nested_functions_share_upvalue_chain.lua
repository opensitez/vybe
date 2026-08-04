-- vybe-test: lua/scoping/scoping_nested_functions_share_upvalue_chain
-- origin: languages/lua/tests/lua/test_scoping.rs

local __w1 = "15"
local __i = 0

local val = 5
local function f1()
  local function f2()
    local function f3()
      val = val + 10
    end
    f3()
  end
  f2()
end
f1()
do local __t = tostring(val); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
