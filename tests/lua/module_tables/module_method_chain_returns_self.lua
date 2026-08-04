-- vybe-test: lua/module_tables/module_method_chain_returns_self
-- origin: languages/lua/tests/lua/test_module_tables.rs

local __w1 = "a-b-c"
local __i = 0

local Builder = {parts = {}}
function Builder:add(s) self.parts[#self.parts+1] = s; return self end
function Builder:build() return table.concat(self.parts, '-') end
do local __t = tostring(Builder:add('a'):add('b'):add('c'):build()); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
