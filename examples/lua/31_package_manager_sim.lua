-- App 31: Package Dependency Resolver (topological sort)

local deps = {
  app = {"http", "auth", "db"},
  http = {"net"},
  auth = {"crypto", "db"},
  db = {"net"},
  crypto = {},
  net = {},
}

local temp, perm, order = {}, {}, {}

local function visit(node)
  if perm[node] then return end
  if temp[node] then error("cycle at " .. node) end
  temp[node] = true
  for _, d in ipairs(deps[node] or {}) do visit(d) end
  temp[node] = nil
  perm[node] = true
  order[#order + 1] = node
end

visit("app")

print("install order:")
for i, p in ipairs(order) do
  print(i, p)
end
