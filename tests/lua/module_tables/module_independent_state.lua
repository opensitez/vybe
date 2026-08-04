-- vybe-test: lua/module_tables/module_independent_state
-- origin: languages/lua/tests/lua/test_module_tables.rs

local __w1 = "2,1"
local __i = 0

local function new_stack()
  local data = {}
  return {
    push = function(v) data[#data+1] = v end,
    pop = function() local v = data[#data]; data[#data] = nil; return v end,
    size = function() return #data end,
  }
end
local a = new_stack()
local b = new_stack()
a.push(1); a.push(2)
b.push(9)
do local __t = tostring(a.size() .. "," .. b.size()); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
