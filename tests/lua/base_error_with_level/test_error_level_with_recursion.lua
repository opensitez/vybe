-- vybe-test: lua/base_error_with_level/test_error_level_with_recursion
-- origin: languages/lua/tests/lua/test_base_error_with_level.rs

local __w1 = "true"
local __i = 0

local function walk(n)
  if n > 0 then return walk(n - 1) end
  error("deep", 4)
end
local ok, err = pcall(function() walk(1) end)
do local __t = tostring(string.find(err, "deep") ~= nil); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
