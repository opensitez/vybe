-- vybe-test: lua/weak_tables_advanced/memoize_weak_kv
-- origin: languages/lua/tests/lua/test_weak_tables_advanced.rs

local __w1 = "true"
local __i = 0

local cache = setmetatable({}, {__mode="kv"})
local function get_cached(k)
  if not cache[k] then cache[k] = {val = k.val * 2} end
  return cache[k]
end
local key = {val = 10}
local v1 = get_cached(key)
local v2 = get_cached(key)
do local __t = tostring(v1.val == v2.val); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
