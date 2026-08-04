-- vybe-test: lua/closures/closure_in_table_as_method_with_captured_state
-- origin: languages/lua/tests/lua/test_closures.rs

local __w1 = "99"
local __i = 0

local function make_obj(init)
  local val = init
  return {
    get = function() return val end,
    set = function(v) val = v end,
  }
end
local obj = make_obj(10)
obj.set(99)
do local __t = tostring(obj.get()); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
