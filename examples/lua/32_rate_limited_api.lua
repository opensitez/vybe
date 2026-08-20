-- App 32: Rate-Limited API Gateway (token bucket)

local bucket = {
  capacity = 5,
  tokens = 5,
  refill_rate = 1,
  last_time = 0,
}

local function refill(now)
  local elapsed = now - bucket.last_time
  if elapsed > 0 then
    bucket.tokens = math.min(bucket.capacity, bucket.tokens + elapsed * bucket.refill_rate)
    bucket.last_time = now
  end
end

local function allow(now)
  refill(now)
  if bucket.tokens >= 1 then
    bucket.tokens = bucket.tokens - 1
    return true
  end
  return false
end

local requests = {
  {t = 0, path = "/search"}, {t = 0, path = "/search"}, {t = 0, path = "/search"},
  {t = 1, path = "/search"}, {t = 1, path = "/search"}, {t = 1, path = "/search"},
  {t = 3, path = "/search"}, {t = 4, path = "/search"},
}

for _, r in ipairs(requests) do
  print(r.t, r.path, allow(r.t) and "200" or "429")
end
