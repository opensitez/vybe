-- vybe-test: lua/base_tostring_table/test_tostring_table_rounded
-- origin: languages/lua/tests/lua/test_base_tostring_table.rs

local __w1 = "true"
local __i = 0

local t = setmetatable({value = 8}, {__tostring = function(self) return "tbl_" .. tostring(self.value) end}); do local __t = tostring(tostring(t) == "tbl_8"); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
