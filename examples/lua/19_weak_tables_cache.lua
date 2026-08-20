-- 19_weak_tables_cache.lua
-- Demonstrates weak-value cache tables and memoized object wrappers.

local cache = setmetatable({}, {__mode = "v"})

local function get_wrapper(id)
  local obj = cache[id]
  if obj then
    return obj, true
  end

  obj = {
    id = id,
    created = os.time(),
    describe = function(self)
      return string.format("Wrapper<%s> created=%d", self.id, self.created)
    end
  }

  cache[id] = obj
  return obj, false
end

local a1, hit1 = get_wrapper("alpha")
local a2, hit2 = get_wrapper("alpha")
local b1, hit3 = get_wrapper("beta")

print(a1:describe(), "cache hit:", hit1)
print(a2:describe(), "cache hit:", hit2)
print(b1:describe(), "cache hit:", hit3)
print("same alpha object:", rawequal(a1, a2))
