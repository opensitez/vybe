-- vybe-test: lua/load/load_with_reader_function
-- origin: languages/lua/tests/lua/test_load.rs

local __w1 = "99"
local __i = 0

local parts = {"return ", "99", nil}
local i = 0
local f = load(function() i = i + 1; return parts[i] end)
do local __t = tostring(f()); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
