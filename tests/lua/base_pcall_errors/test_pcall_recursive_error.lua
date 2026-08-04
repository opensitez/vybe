-- vybe-test: lua/base_pcall_errors/test_pcall_recursive_error
-- origin: languages/lua/tests/lua/test_base_pcall_errors.rs

local __w1 = "true"
local __i = 0

local function go(n) if n == 0 then error("end") else return go(n-1) end end
local ok, err = pcall(function() go(2) end)
do local __t = tostring(ok == false and string.find(err, "end") ~= nil); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
