-- Full App 03: Route Planner (Dijkstra shortest path)

local graph = {
  A = {B = 4, C = 2},
  B = {A = 4, C = 1, D = 5},
  C = {A = 2, B = 1, D = 8, E = 10},
  D = {B = 5, C = 8, E = 2, F = 6},
  E = {C = 10, D = 2, F = 3},
  F = {D = 6, E = 3},
}

local function dijkstra(start)
  local dist, prev, unvisited = {}, {}, {}
  for node, _ in pairs(graph) do
    dist[node] = math.huge
    unvisited[node] = true
  end
  dist[start] = 0

  while true do
    local u, best = nil, math.huge
    for n, _ in pairs(unvisited) do
      if dist[n] < best then u, best = n, dist[n] end
    end
    if not u then break end
    unvisited[u] = nil
    for v, w in pairs(graph[u]) do
      local alt = dist[u] + w
      if alt < dist[v] then
        dist[v] = alt
        prev[v] = u
      end
    end
  end
  return dist, prev
end

local function build_path(prev, target)
  local path = {target}
  local cur = target
  while prev[cur] do
    cur = prev[cur]
    table.insert(path, 1, cur)
  end
  return path
end

local dist, prev = dijkstra("A")
local target = "F"
local path = build_path(prev, target)
print("best path:", table.concat(path, " -> "))
print("total cost:", dist[target])
