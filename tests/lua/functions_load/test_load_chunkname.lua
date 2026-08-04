-- vybe-test: lua/functions_load/test_load_chunkname
-- origin: languages/lua/tests/lua/test_functions_load.rs

local __w1 = "true"
local __i = 0

local f, err = load('error()', 'mychunk'); local ok, err_msg = pcall(f); do local __t = tostring(tostring(string.find(err_msg, 'mychunk') ~= nil)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
