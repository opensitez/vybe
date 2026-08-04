-- vybe-test: lua/metatables_index/test_index_method_call_from_base
-- origin: languages/lua/tests/lua/test_metatables_index.rs

local __w1 = "20"
local __i = 0

local base={v=10}; function base:get() return self.v end; local child=setmetatable({v=20}, {__index=base}); do local __t = tostring(child:get()); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
