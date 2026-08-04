-- vybe-test: lua/programs/metatable_index_fallback_chain
-- origin: languages/lua/tests/lua/test_programs.rs

local __w1 = "lua"
local __i = 0

local defaults = {lang = "lua"}
local t = setmetatable({}, {__index = defaults})
do local __t = tostring(t.lang); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
