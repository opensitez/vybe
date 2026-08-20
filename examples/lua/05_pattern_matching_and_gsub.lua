-- App 05: Access Log Analyzer (parse, aggregate, anonymize)

local lines = {
  '10.0.0.8 - - [23/Jun/2026:10:00:00] "GET /api/users HTTP/1.1" 200 512',
  '10.0.0.9 - - [23/Jun/2026:10:00:01] "POST /api/orders HTTP/1.1" 201 128',
  '10.0.0.8 - - [23/Jun/2026:10:00:02] "GET /api/users HTTP/1.1" 500 64',
}

local by_status, by_path = {}, {}

for _, line in ipairs(lines) do
  local ip, method, path, status, size = line:match('^(%S+).+"(%u+) (%S+) HTTP/%d%.%d" (%d+) (%d+)$')
  status, size = tonumber(status), tonumber(size)
  by_status[status] = (by_status[status] or 0) + 1
  by_path[path] = (by_path[path] or {hits = 0, bytes = 0})
  by_path[path].hits = by_path[path].hits + 1
  by_path[path].bytes = by_path[path].bytes + size

  local masked = line:gsub("^(%d+)%.(%d+)%.(%d+)%.(%d+)", "%1.%2.x.x")
  print("anonymized:", masked)
end

print("status summary")
for code, count in pairs(by_status) do
  print(code, count)
end

print("path summary")
for path, stats in pairs(by_path) do
  print(path, "hits=" .. stats.hits, "bytes=" .. stats.bytes)
end
