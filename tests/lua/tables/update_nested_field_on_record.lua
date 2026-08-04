-- vybe-test: lua/tables/update_nested_field_on_record
-- origin: languages/lua/tests/lua/test_tables.rs

local __w1 = "bob"
local __i = 0

local user = {profile = {name = "ada"}}
user.profile.name = "bob"
do local __t = tostring(user.profile.name); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
