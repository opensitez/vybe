-- vybe-test: lua/tables/hash_key_bracket_syntax_for_non_identifier
-- origin: languages/lua/tests/lua/test_tables.rs

local __w1 = "7"
local __i = 0

local t = {}
t["key-name"] = 7
do local __t = tostring(t["key-name"]); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
