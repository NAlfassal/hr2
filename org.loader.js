// org.loader.js
const allowedChild = {
  lv1: ['lv2'],
  lv2: ['lv3'],
  lv3: ['lv4'],
  lv4: []
};

function canNest(parent, childType) {
  if (!parent) return childType === 'lv1';
  return (allowedChild[parent.type] || []).includes(childType);
}

function slugId(prefix, name) {
  const s = (name || '')
    .toLowerCase()
    .replace(/\s+/g, '-')
    .replace(/[^\p{Letter}\p{Number}\-]/gu, '');
  return `${prefix}-${s}-${Math.random().toString(36).slice(2,7)}`;
}


function addNode(store, partial) {
  const parent = partial.parentId ? store.nodes[partial.parentId] : undefined;
  if (!canNest(parent, partial.type)) {
    throw new Error(`Invalid placement: ${partial.type} under ${parent?.type || 'ROOT'}`);
  }
  const node = { ...partial, children: partial.children || [] };
  store.nodes[node.id] = node;
  if (parent) parent.children.push(node.id);
  return node.id;
}

function buildStoreFromJson(json) {
  const store = { nodes: {} };

  function dfs(raw, parentId) {
    const id = slugId(raw.type, raw.name);
    addNode(store, { id, name: raw.name, type: raw.type, parentId });
    (raw.children || []).forEach(ch => dfs(ch, id));
  }

  (json.units || []).forEach(u => dfs(u, undefined));
  return store;
}

// ---------- Utilities ----------
function listChildren(store, parentId, typeFilter) {
  const items = Object.values(store.nodes).filter(n => n.parentId === parentId);
  return typeFilter ? items.filter(n => n.type === typeFilter) : items;
}

function listByType(store, type) {
  return Object.values(store.nodes).filter(n => n.type === type);
}

function getPath(store, nodeId) {
  const path = [];
  let cur = store.nodes[nodeId];
  while (cur && cur.parentId) {
    cur = store.nodes[cur.parentId];
    if (cur) path.unshift(cur);
  }
  return path; // ancestors only (not including the node)
}

function validate(store) {
  const errs = [];
  for (const n of Object.values(store.nodes)) {
    const parent = n.parentId ? store.nodes[n.parentId] : undefined;
    if (!canNest(parent, n.type)) errs.push(`Invalid parent for ${n.name} (${n.type})`);
    for (const cid of n.children) {
      const ch = store.nodes[cid];
      if (!ch) { errs.push(`Missing child ${cid} under ${n.name}`); continue; }
      if (ch.parentId !== n.id) errs.push(`Parent/child mismatch: ${n.name} ↔ ${ch.name}`);
      if (!allowedChild[n.type].includes(ch.type)) {
        errs.push(`Not allowed: ${ch.type} under ${n.type} (${ch.name} under ${n.name})`);
      }
    }
  }
  return errs;
}

function moveNode(store, nodeId, newParentId) {
  const node = store.nodes[nodeId];
  const newParent = newParentId ? store.nodes[newParentId] : undefined;
  if (!node) throw new Error('Node not found');
  if (!canNest(newParent, node.type)) throw new Error('Invalid move');

  // detach
  if (node.parentId) {
    const oldParent = store.nodes[node.parentId];
    oldParent.children = oldParent.children.filter(id => id !== nodeId);
  }
  // attach
  node.parentId = newParentId;
  if (newParent) newParent.children.push(nodeId);
}

// Helpful finder
function findByName(store, name, type) {
  return Object.values(store.nodes).find(n => n.name === name && (!type || n.type === type));
}

export {
  buildStoreFromJson,
  listChildren,
  listByType,
  getPath,
  validate,
  moveNode,
  findByName
};

