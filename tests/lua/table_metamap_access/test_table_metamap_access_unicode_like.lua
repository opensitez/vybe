-- vybe-test: lua/table_metamap_access/test_table_metamap_access_unicode_like
-- origin: languages/lua/tests/lua/test_table_metamap_access.rs

local __w1 = "true"
local __i = 0

local mt = {__index = function(_, k) return k end}
local t = setmetatable({}, mt)
do local __t = tostring(t["k20"] == "k20"); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
