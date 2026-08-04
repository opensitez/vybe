-- vybe-test: lua/load/load_with_invalid_binary_chunk_detection
-- origin: languages/lua/tests/lua/test_load.rs

local __w1 = "true"
local __i = 0

local f, err = load("\27Lua\000")
do local __t = tostring(f == nil and type(err) == "string"); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
