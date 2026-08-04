-- vybe-test: lua/base_setmetatable_guarded/newindex_guard_rejects_missing_key_assignment
-- origin: languages/lua/tests/lua/test_base_setmetatable_guarded.rs

local __w1 = "false"
local __i = 0

local t = {}
setmetatable(t, {__newindex = function() error("guard") end})
local ok = pcall(function() t.x = 1 end)
do local __t = tostring(ok); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
