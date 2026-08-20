-- App 04: Config Loader with validation and fallback

local raw_config = {
  '{"port":8080,"mode":"prod","workers":4}',
  '{"port":-1,"mode":"prod","workers":4}',
  '{"port":3000,"mode":"dev","workers":0}',
}

local function parse_jsonish(s)
  local cfg = {}
  for k, v in s:gmatch('"([%w_]+)":([%w%-]+)') do
    cfg[k] = tonumber(v) or v
  end
  return cfg
end

local function validate(cfg)
  assert(type(cfg.port) == "number" and cfg.port > 0 and cfg.port < 65536, "invalid port")
  assert(cfg.mode == "dev" or cfg.mode == "prod", "invalid mode")
  assert(type(cfg.workers) == "number" and cfg.workers >= 1, "invalid workers")
  return cfg
end

local function load_one(line)
  return validate(parse_jsonish(line))
end

for i, line in ipairs(raw_config) do
  local ok, result = xpcall(function() return load_one(line) end, function(e)
    return "config #" .. i .. " failed: " .. tostring(e)
  end)
  if ok then
    print(string.format("config #%d loaded: port=%d mode=%s workers=%d", i, result.port, result.mode, result.workers))
  else
    print(result)
  end
end
