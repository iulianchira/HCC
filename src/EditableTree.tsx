import { useEffect, useMemo, useRef, useState } from "react";
import type { DragEvent, FormEvent } from "react";
import { Button } from "@fluentui/react-button";
import { Checkbox } from "@fluentui/react-checkbox";
import { ColorArea, ColorPicker, ColorSlider } from "@fluentui/react-color-picker";
import {
  DrawerBody,
  DrawerFooter,
  DrawerHeader,
  DrawerHeaderTitle,
  OverlayDrawer
} from "@fluentui/react-drawer";
import {
  AddRegular,
  ArrowCollapseAllRegular,
  ArrowDownRegular,
  ArrowExpandAllRegular,
  ArrowUpRegular,
  CheckmarkRegular,
  DeleteRegular,
  DismissRegular,
  EditRegular,
  MoreHorizontalRegular,
  ReOrderDotsVerticalRegular,
  SearchRegular,
  TextIndentDecreaseRegular,
  TextIndentIncreaseRegular
} from "@fluentui/react-icons";
import { Menu, MenuItem, MenuList, MenuPopover, MenuTrigger } from "@fluentui/react-menu";
import { Input } from "@fluentui/react-input";
import { Text } from "@fluentui/react-text";
import { Tooltip } from "@fluentui/react-tooltip";
import { Tree, TreeItem, TreeItemLayout, type TreeItemValue, type TreeOpenChangeData } from "@fluentui/react-tree";
import type { TreeControlMode, TreeNode, TreeSelectionBehavior, TreeSelectionMode } from "./types";
import "./EditableTree.css";

type EditableTreeProps = {
  initialNodes: TreeNode[];
  mode: TreeControlMode;
  dragAndDropEnabled?: boolean;
  selectionMode: TreeSelectionMode;
  selectionBehavior: TreeSelectionBehavior;
  selectedItemIds: TreeItemValue[];
  restrictSelectionToLeafNodes?: boolean;
  onSelectedItemIdsChange?: (selectedItemIds: TreeItemValue[]) => void;
  onChange?: (nodes: TreeNode[]) => void;
};

type HsvColor = {
  h: number;
  s: number;
  v: number;
  a?: number;
};

type NodePath = number[];
type DropPosition = "before" | "inside" | "after";
type CascadeNodeState = "checked" | "mixed" | "unchecked";
type SearchFilterState = {
  visibleNodeIds: Set<TreeItemValue>;
  matchingNodeIds: Set<TreeItemValue>;
  branchNodeIdsToOpen: Set<TreeItemValue>;
};

const DEFAULT_TEXT_COLOR = "#FFFFFF";
const DEFAULT_BACKGROUND_COLOR = "#0F6CBD";
const TOOLTIP_SHOW_DELAY_MS = 2000;
const DRAG_NODE_MIME_TYPE = "application/x-editable-tree-node-id";
const FULL_ACTION_BUTTON_WIDTH_PX = 28;
const FULL_ACTION_BUTTON_GAP_PX = 4;
const FULL_ACTION_COUNT = 7;
const MIN_LABEL_WIDTH_PX = 180;
const FULL_ACTIONS_REQUIRED_WIDTH_PX =
  FULL_ACTION_COUNT * FULL_ACTION_BUTTON_WIDTH_PX + (FULL_ACTION_COUNT - 1) * FULL_ACTION_BUTTON_GAP_PX + MIN_LABEL_WIDTH_PX;

const clamp = (value: number, min: number, max: number): number => Math.min(Math.max(value, min), max);

const componentToHex = (value: number): string => value.toString(16).padStart(2, "0").toUpperCase();

const hsvToHex = (color: HsvColor): string => {
  const hue = ((color.h % 360) + 360) % 360;
  const saturation = clamp(color.s, 0, 1);
  const value = clamp(color.v, 0, 1);
  const chroma = value * saturation;
  const x = chroma * (1 - Math.abs(((hue / 60) % 2) - 1));
  const match = value - chroma;

  let red = 0;
  let green = 0;
  let blue = 0;

  if (hue < 60) {
    red = chroma;
    green = x;
  } else if (hue < 120) {
    red = x;
    green = chroma;
  } else if (hue < 180) {
    green = chroma;
    blue = x;
  } else if (hue < 240) {
    green = x;
    blue = chroma;
  } else if (hue < 300) {
    red = x;
    blue = chroma;
  } else {
    red = chroma;
    blue = x;
  }

  const r = Math.round((red + match) * 255);
  const g = Math.round((green + match) * 255);
  const b = Math.round((blue + match) * 255);

  return `#${componentToHex(r)}${componentToHex(g)}${componentToHex(b)}`;
};

const hexToHsv = (value: string): HsvColor => {
  const normalized = value.trim().replace("#", "");
  const hex = normalized.length === 3 ? normalized.split("").map((piece) => piece + piece).join("") : normalized;

  if (!/^[0-9a-fA-F]{6}$/.test(hex)) {
    return { h: 207, s: 0.92, v: 0.74, a: 1 };
  }

  const red = parseInt(hex.slice(0, 2), 16) / 255;
  const green = parseInt(hex.slice(2, 4), 16) / 255;
  const blue = parseInt(hex.slice(4, 6), 16) / 255;

  const max = Math.max(red, green, blue);
  const min = Math.min(red, green, blue);
  const diff = max - min;

  let hue = 0;
  if (diff !== 0) {
    if (max === red) {
      hue = ((green - blue) / diff) % 6;
    } else if (max === green) {
      hue = (blue - red) / diff + 2;
    } else {
      hue = (red - green) / diff + 4;
    }
    hue *= 60;
    if (hue < 0) {
      hue += 360;
    }
  }

  const saturation = max === 0 ? 0 : diff / max;
  return { h: hue, s: saturation, v: max, a: 1 };
};

const DEFAULT_TEXT_HSV_COLOR = hexToHsv(DEFAULT_TEXT_COLOR);
const DEFAULT_BACKGROUND_HSV_COLOR = hexToHsv(DEFAULT_BACKGROUND_COLOR);

const cloneNodes = (nodes: TreeNode[]): TreeNode[] =>
  nodes.map((node) => ({
    ...node,
    children: node.children ? cloneNodes(node.children) : undefined
  }));

const createNode = (
  label = "New item",
  value = "new-item",
  textColor = DEFAULT_TEXT_COLOR,
  backgroundColor = DEFAULT_BACKGROUND_COLOR
): TreeNode => ({
  id: globalThis.crypto?.randomUUID?.() ?? `${Date.now()}-${Math.random().toString(16).slice(2)}`,
  label,
  value,
  textColor,
  backgroundColor
});

const countNodes = (nodes: TreeNode[]): number =>
  nodes.reduce((total, node) => total + 1 + countNodes(node.children ?? []), 0);

const collectNodeIds = (nodes: TreeNode[]): TreeItemValue[] =>
  nodes.flatMap((node) => [node.id, ...collectNodeIds(node.children ?? [])]);

const collectLeafNodeIds = (nodes: TreeNode[]): TreeItemValue[] =>
  nodes.flatMap((node) =>
    node.children?.length ? collectLeafNodeIds(node.children) : [node.id]
  );

// Collect IDs toggled together in cascade mode.
const collectSelectableSubtreeIds = (
  node: TreeNode,
  restrictSelectionToLeafNodes: boolean,
  includeSelf = true
): TreeItemValue[] => {
  const hasChildren = (node.children?.length ?? 0) > 0;
  const isSelectable = !restrictSelectionToLeafNodes || !hasChildren;
  const ownId = includeSelf && isSelectable ? [node.id] : [];

  return [
    ...ownId,
    ...(node.children ?? []).flatMap((child) => collectSelectableSubtreeIds(child, restrictSelectionToLeafNodes, true))
  ];
};

// Compute tri-state checkbox value for each node from its descendants.
const getCascadeNodeState = (
  node: TreeNode,
  selectedSet: ReadonlySet<TreeItemValue>,
  restrictSelectionToLeafNodes: boolean,
  checkStateById?: Map<TreeItemValue, boolean | "mixed">
): CascadeNodeState => {
  const children = node.children ?? [];
  const hasChildren = children.length > 0;
  const isSelectable = !restrictSelectionToLeafNodes || !hasChildren;

  if (!hasChildren) {
    const isChecked = isSelectable && selectedSet.has(node.id);
    checkStateById?.set(node.id, isChecked);
    return isChecked ? "checked" : "unchecked";
  }

  let allChildrenChecked = true;
  let hasCheckedOrMixedChild = false;

  for (const child of children) {
    const childState = getCascadeNodeState(child, selectedSet, restrictSelectionToLeafNodes, checkStateById);
    if (childState !== "checked") {
      allChildrenChecked = false;
    }
    if (childState !== "unchecked") {
      hasCheckedOrMixedChild = true;
    }
  }

  // Parents are checked only when all descendants are checked.
  const nodeState: CascadeNodeState = allChildrenChecked ? "checked" : hasCheckedOrMixedChild ? "mixed" : "unchecked";
  if (isSelectable) {
    checkStateById?.set(node.id, nodeState === "checked" ? true : nodeState === "mixed" ? "mixed" : false);
  } else {
    checkStateById?.set(node.id, nodeState === "mixed" ? "mixed" : nodeState === "checked");
  }

  return nodeState;
};

// Canonicalize cascade selection so parent IDs follow descendant state.
const normalizeCascadeSelection = (
  nodes: TreeNode[],
  selectedSet: ReadonlySet<TreeItemValue>,
  restrictSelectionToLeafNodes: boolean
): Set<TreeItemValue> => {
  const normalizedSelectedItems = new Set(selectedSet);

  const visit = (node: TreeNode): CascadeNodeState => {
    const children = node.children ?? [];
    const hasChildren = children.length > 0;
    const isSelectable = !restrictSelectionToLeafNodes || !hasChildren;

    if (!hasChildren) {
      const isChecked = isSelectable && normalizedSelectedItems.has(node.id);
      if (isSelectable && isChecked) {
        normalizedSelectedItems.add(node.id);
      } else {
        normalizedSelectedItems.delete(node.id);
      }
      return isChecked ? "checked" : "unchecked";
    }

    let allChildrenChecked = true;
    let hasCheckedOrMixedChild = false;

    for (const child of children) {
      const childState = visit(child);
      if (childState !== "checked") {
        allChildrenChecked = false;
      }
      if (childState !== "unchecked") {
        hasCheckedOrMixedChild = true;
      }
    }

    const nodeState: CascadeNodeState = allChildrenChecked ? "checked" : hasCheckedOrMixedChild ? "mixed" : "unchecked";
    if (!isSelectable) {
      // Leaf-only mode keeps branch IDs out of the selected IDs list.
      normalizedSelectedItems.delete(node.id);
      return nodeState;
    }

    if (nodeState === "checked") {
      normalizedSelectedItems.add(node.id);
    } else {
      normalizedSelectedItems.delete(node.id);
    }

    return nodeState;
  };

  for (const rootNode of nodes) {
    visit(rootNode);
  }

  return normalizedSelectedItems;
};

const isSameSelection = (left: TreeItemValue[], right: ReadonlySet<TreeItemValue>): boolean => {
  if (left.length !== right.size) {
    return false;
  }

  return left.every((value) => right.has(value));
};

// Keep matched nodes plus ancestors visible, and open ancestor branches while searching.
const buildSearchFilterState = (nodes: TreeNode[], normalizedQuery: string): SearchFilterState => {
  const visibleNodeIds = new Set<TreeItemValue>();
  const matchingNodeIds = new Set<TreeItemValue>();
  const branchNodeIdsToOpen = new Set<TreeItemValue>();

  if (!normalizedQuery) {
    return { visibleNodeIds, matchingNodeIds, branchNodeIdsToOpen };
  }

  const visitNode = (node: TreeNode): boolean => {
    const ownMatch = node.label.toLowerCase().includes(normalizedQuery);
    let hasMatchingDescendant = false;

    for (const child of node.children ?? []) {
      if (visitNode(child)) {
        hasMatchingDescendant = true;
      }
    }

    const isVisible = ownMatch || hasMatchingDescendant;
    if (!isVisible) {
      return false;
    }

    visibleNodeIds.add(node.id);
    if (ownMatch) {
      matchingNodeIds.add(node.id);
    }
    if (hasMatchingDescendant) {
      branchNodeIdsToOpen.add(node.id);
    }

    return true;
  };

  for (const rootNode of nodes) {
    visitNode(rootNode);
  }

  return { visibleNodeIds, matchingNodeIds, branchNodeIdsToOpen };
};

const hasNode = (nodes: TreeNode[], nodeId: string): boolean =>
  nodes.some((node) => node.id === nodeId || hasNode(node.children ?? [], nodeId));

const getNodeById = (nodes: TreeNode[], nodeId: string): TreeNode | undefined => {
  for (const node of nodes) {
    if (node.id === nodeId) {
      return node;
    }

    const match = getNodeById(node.children ?? [], nodeId);
    if (match) {
      return match;
    }
  }

  return undefined;
};

const collectBranchIds = (nodes: TreeNode[]): TreeItemValue[] =>
  nodes.flatMap((node) =>
    node.children?.length ? [node.id, ...collectBranchIds(node.children)] : collectBranchIds(node.children ?? [])
  );

const getPathById = (nodes: TreeNode[], nodeId: string, parentPath: NodePath = []): NodePath | undefined => {
  for (let index = 0; index < nodes.length; index += 1) {
    const node = nodes[index];
    const path = [...parentPath, index];

    if (node.id === nodeId) {
      return path;
    }

    const childPath = getPathById(node.children ?? [], nodeId, path);
    if (childPath) {
      return childPath;
    }
  }

  return undefined;
};

const getPathNodeIds = (nodes: TreeNode[], path: NodePath): TreeItemValue[] => {
  const ids: TreeItemValue[] = [];
  let level: TreeNode[] = nodes;

  for (const index of path) {
    const node = level[index];
    if (!node) {
      break;
    }

    ids.push(node.id);
    level = node.children ?? [];
  }

  return ids;
};

const getNodePathById = (
  nodes: TreeNode[],
  nodeId: string,
  parentPath: TreeNode[] = []
): TreeNode[] | undefined => {
  for (const node of nodes) {
    const currentPath = [...parentPath, node];
    if (node.id === nodeId) {
      return currentPath;
    }

    const childPath = getNodePathById(node.children ?? [], nodeId, currentPath);
    if (childPath) {
      return childPath;
    }
  }

  return undefined;
};

const getNodeAtPath = (nodes: TreeNode[], path: NodePath): TreeNode | undefined => {
  let level: TreeNode[] = nodes;
  let current: TreeNode | undefined;

  for (const index of path) {
    current = level[index];
    if (!current) {
      return undefined;
    }

    level = current.children ?? [];
  }

  return current;
};

const getSiblingList = (nodes: TreeNode[], parentPath: NodePath): TreeNode[] | undefined => {
  if (parentPath.length === 0) {
    return nodes;
  }

  const parent = getNodeAtPath(nodes, parentPath);
  return parent?.children;
};

const updateNodeById = (
  nodes: TreeNode[],
  nodeId: string,
  updater: (node: TreeNode) => TreeNode
): TreeNode[] =>
  nodes.map((node) => {
    if (node.id === nodeId) {
      return updater(node);
    }

    if (!node.children?.length) {
      return node;
    }

    return {
      ...node,
      children: updateNodeById(node.children, nodeId, updater)
    };
  });

const addChildById = (nodes: TreeNode[], parentId: string, child: TreeNode): TreeNode[] =>
  updateNodeById(nodes, parentId, (node) => ({
    ...node,
    children: [...(node.children ?? []), child]
  }));

const updateNodeDetailsById = (
  nodes: TreeNode[],
  nodeId: string,
  nextLabel: string,
  nextValue: string,
  nextTextColor: string,
  nextBackgroundColor: string
): TreeNode[] =>
  updateNodeById(nodes, nodeId, (node) => ({
    ...node,
    label: nextLabel,
    value: nextValue,
    textColor: nextTextColor,
    backgroundColor: nextBackgroundColor
  }));

const removeNodeById = (nodes: TreeNode[], nodeId: string): TreeNode[] =>
  nodes
    .filter((node) => node.id !== nodeId)
    .map((node) => {
      if (!node.children?.length) {
        return node;
      }

      const children = removeNodeById(node.children, nodeId);
      return children.length > 0 ? { ...node, children } : { ...node, children: undefined };
    });

const moveNodeUpByPath = (nodes: TreeNode[], path: NodePath): TreeNode[] => {
  const index = path[path.length - 1];
  if (index === undefined || index <= 0) {
    return nodes;
  }

  const nextNodes = cloneNodes(nodes);
  const parentPath = path.slice(0, -1);
  const siblings = getSiblingList(nextNodes, parentPath);
  if (!siblings || index >= siblings.length) {
    return nodes;
  }

  [siblings[index - 1], siblings[index]] = [siblings[index], siblings[index - 1]];
  return nextNodes;
};

const moveNodeDownByPath = (nodes: TreeNode[], path: NodePath): TreeNode[] => {
  const index = path[path.length - 1];
  if (index === undefined) {
    return nodes;
  }

  const nextNodes = cloneNodes(nodes);
  const parentPath = path.slice(0, -1);
  const siblings = getSiblingList(nextNodes, parentPath);
  if (!siblings || index < 0 || index >= siblings.length - 1) {
    return nodes;
  }

  [siblings[index], siblings[index + 1]] = [siblings[index + 1], siblings[index]];
  return nextNodes;
};

// Move a node under its previous sibling as the last child.
const indentNodeByPath = (nodes: TreeNode[], path: NodePath): TreeNode[] => {
  const index = path[path.length - 1];
  if (index === undefined || index <= 0) {
    return nodes;
  }

  const nextNodes = cloneNodes(nodes);
  const parentPath = path.slice(0, -1);
  const siblings = getSiblingList(nextNodes, parentPath);
  if (!siblings || index >= siblings.length) {
    return nodes;
  }

  const [movingNode] = siblings.splice(index, 1);
  const previousSibling = siblings[index - 1];
  previousSibling.children = [...(previousSibling.children ?? []), movingNode];

  if (parentPath.length > 0 && siblings.length === 0) {
    const parentNode = getNodeAtPath(nextNodes, parentPath);
    if (parentNode) {
      parentNode.children = undefined;
    }
  }

  return nextNodes;
};

// Lift a node one level up and place it right after its parent.
const outdentNodeByPath = (nodes: TreeNode[], path: NodePath): TreeNode[] => {
  const index = path[path.length - 1];
  const parentPath = path.slice(0, -1);
  if (index === undefined || parentPath.length === 0) {
    return nodes;
  }

  const parentIndex = parentPath[parentPath.length - 1];
  if (parentIndex === undefined) {
    return nodes;
  }

  const nextNodes = cloneNodes(nodes);
  const currentSiblings = getSiblingList(nextNodes, parentPath);
  const grandParentPath = parentPath.slice(0, -1);
  const grandSiblings = getSiblingList(nextNodes, grandParentPath);

  if (!currentSiblings || !grandSiblings || index < 0 || index >= currentSiblings.length) {
    return nodes;
  }

  const [movingNode] = currentSiblings.splice(index, 1);
  grandSiblings.splice(parentIndex + 1, 0, movingNode);

  const parentNode = getNodeAtPath(nextNodes, parentPath);
  if (parentNode && parentNode.children && parentNode.children.length === 0) {
    parentNode.children = undefined;
  }

  return nextNodes;
};

const isPathPrefix = (prefixPath: NodePath, path: NodePath): boolean =>
  prefixPath.length <= path.length && prefixPath.every((segment, index) => path[index] === segment);

const removeNodeAtPath = (nodes: TreeNode[], path: NodePath): { nextNodes: TreeNode[]; removedNode?: TreeNode } => {
  const index = path[path.length - 1];
  if (index === undefined) {
    return { nextNodes: nodes };
  }

  const nextNodes = cloneNodes(nodes);
  const parentPath = path.slice(0, -1);
  const siblings = getSiblingList(nextNodes, parentPath);
  if (!siblings || index < 0 || index >= siblings.length) {
    return { nextNodes: nodes };
  }

  const [removedNode] = siblings.splice(index, 1);
  if (!removedNode) {
    return { nextNodes: nodes };
  }

  if (parentPath.length > 0 && siblings.length === 0) {
    const parentNode = getNodeAtPath(nextNodes, parentPath);
    if (parentNode) {
      parentNode.children = undefined;
    }
  }

  return { nextNodes, removedNode };
};

const insertNodeAtPath = (nodes: TreeNode[], targetPath: NodePath, movingNode: TreeNode, position: DropPosition): TreeNode[] => {
  if (position === "inside") {
    const targetNode = getNodeAtPath(nodes, targetPath);
    if (!targetNode) {
      return nodes;
    }

    targetNode.children = [...(targetNode.children ?? []), movingNode];
    return nodes;
  }

  const targetIndex = targetPath[targetPath.length - 1];
  if (targetIndex === undefined) {
    return nodes;
  }

  const parentPath = targetPath.slice(0, -1);
  const siblings = getSiblingList(nodes, parentPath);
  if (!siblings) {
    return nodes;
  }

  const insertionIndex = position === "before" ? targetIndex : targetIndex + 1;
  siblings.splice(insertionIndex, 0, movingNode);
  return nodes;
};

const moveNodeByDrop = (nodes: TreeNode[], sourceId: string, targetId: string, position: DropPosition): TreeNode[] => {
  if (sourceId === targetId) {
    return nodes;
  }

  const sourcePath = getPathById(nodes, sourceId);
  const targetPath = getPathById(nodes, targetId);
  if (!sourcePath || !targetPath) {
    return nodes;
  }

  // Block drops into the dragged node subtree.
  if (isPathPrefix(sourcePath, targetPath)) {
    return nodes;
  }

  const { nextNodes, removedNode } = removeNodeAtPath(nodes, sourcePath);
  if (!removedNode) {
    return nodes;
  }

  const nextTargetPath = getPathById(nextNodes, targetId);
  if (!nextTargetPath) {
    return nodes;
  }

  return insertNodeAtPath(nextNodes, nextTargetPath, removedNode, position);
};

function EditableTree({
  initialNodes,
  mode,
  dragAndDropEnabled = true,
  selectionMode,
  selectionBehavior,
  selectedItemIds,
  restrictSelectionToLeafNodes = false,
  onSelectedItemIdsChange,
  onChange
}: EditableTreeProps) {
  const [nodes, setNodes] = useState<TreeNode[]>(initialNodes);
  const [openItems, setOpenItems] = useState<Set<TreeItemValue>>(() => new Set(collectBranchIds(initialNodes)));
  const [drawerOpen, setDrawerOpen] = useState(false);
  const [editingId, setEditingId] = useState<string | null>(null);
  const [pendingNewId, setPendingNewId] = useState<string | null>(null);
  const [searchQuery, setSearchQuery] = useState("");
  const [draftLabel, setDraftLabel] = useState("");
  const [draftValue, setDraftValue] = useState("");
  const [draftTextHsvColor, setDraftTextHsvColor] = useState<HsvColor>(DEFAULT_TEXT_HSV_COLOR);
  const [draftBackgroundHsvColor, setDraftBackgroundHsvColor] = useState<HsvColor>(DEFAULT_BACKGROUND_HSV_COLOR);
  const [draggedNodeId, setDraggedNodeId] = useState<string | null>(null);
  const [dropHint, setDropHint] = useState<{ targetId: string; position: DropPosition } | null>(null);
  const treeRootRef = useRef<HTMLElement | null>(null);
  const [actionsOverflowByNodeId, setActionsOverflowByNodeId] = useState<Record<string, boolean>>({});

  const totalNodes = useMemo(() => countNodes(nodes), [nodes]);
  const branchNodeIds = useMemo(() => collectBranchIds(nodes), [nodes]);
  const draftTextHexColor = useMemo(() => hsvToHex(draftTextHsvColor), [draftTextHsvColor]);
  const draftBackgroundHexColor = useMemo(() => hsvToHex(draftBackgroundHsvColor), [draftBackgroundHsvColor]);
  // Path preview mirrors the in-progress label/color edits for the current node.
  const editingPathNodes = useMemo(() => {
    if (!editingId) {
      return [] as TreeNode[];
    }

    const pathNodes = getNodePathById(nodes, editingId);
    if (!pathNodes?.length) {
      return [] as TreeNode[];
    }

    return pathNodes.map((pathNode, pathIndex) => {
      if (pathIndex !== pathNodes.length - 1) {
        return pathNode;
      }

      return {
        ...pathNode,
        label: draftLabel.trim() || pathNode.label,
        textColor: draftTextHexColor,
        backgroundColor: draftBackgroundHexColor
      };
    });
  }, [draftBackgroundHexColor, draftLabel, draftTextHexColor, editingId, nodes]);
  const selectedItemsSet = useMemo(() => new Set(selectedItemIds), [selectedItemIds]);
  const normalizedSearchQuery = useMemo(() => searchQuery.trim().toLowerCase(), [searchQuery]);
  const isSearchActive = normalizedSearchQuery.length > 0;
  const { visibleNodeIds: searchVisibleNodeIds, matchingNodeIds: searchMatchingNodeIds, branchNodeIdsToOpen } = useMemo(
    () => buildSearchFilterState(nodes, normalizedSearchQuery),
    [nodes, normalizedSearchQuery]
  );
  const isSelectMode = mode === "select";
  const isDragAndDropActive = !isSelectMode && dragAndDropEnabled;
  const isCascadeSelection = isSelectMode && selectionMode === "multiple" && selectionBehavior === "cascade";
  const effectiveOpenItems = useMemo(() => {
    if (!isSearchActive) {
      return openItems;
    }

    // Search expansion is derived state; it does not overwrite user-controlled openItems.
    return new Set([...openItems, ...branchNodeIdsToOpen]);
  }, [branchNodeIdsToOpen, isSearchActive, openItems]);
  const cascadeCheckStateById = useMemo(() => {
    if (!isCascadeSelection) {
      return new Map<TreeItemValue, boolean | "mixed">();
    }

    const selectedSet = new Set(selectedItemIds);
    const checkStateById = new Map<TreeItemValue, boolean | "mixed">();
    for (const rootNode of nodes) {
      getCascadeNodeState(rootNode, selectedSet, restrictSelectionToLeafNodes, checkStateById);
    }

    return checkStateById;
  }, [isCascadeSelection, nodes, restrictSelectionToLeafNodes, selectedItemIds]);

  useEffect(() => {
    onChange?.(nodes);
  }, [nodes, onChange]);

  useEffect(() => {
    if (mode === "select" && drawerOpen) {
      if (pendingNewId) {
        setNodes((prevNodes) => removeNodeById(prevNodes, pendingNewId));
      }

      setDrawerOpen(false);
      setEditingId(null);
      setPendingNewId(null);
      setDraftLabel("");
      setDraftValue("");
      setDraftTextHsvColor(DEFAULT_TEXT_HSV_COLOR);
      setDraftBackgroundHsvColor(DEFAULT_BACKGROUND_HSV_COLOR);
    }
  }, [mode, drawerOpen, pendingNewId]);

  useEffect(() => {
    if (!onSelectedItemIdsChange) {
      return;
    }

    // Remove selections that became invalid after edits or leaf-only mode changes.
    const validNodeIds = new Set(collectNodeIds(nodes));
    const leafNodeIds = new Set(collectLeafNodeIds(nodes));
    const filteredSelectedItems = selectedItemIds.filter(
      (itemId) => validNodeIds.has(itemId) && (!restrictSelectionToLeafNodes || leafNodeIds.has(itemId))
    );
    if (filteredSelectedItems.length !== selectedItemIds.length) {
      onSelectedItemIdsChange(filteredSelectedItems);
    }
  }, [nodes, selectedItemIds, onSelectedItemIdsChange, restrictSelectionToLeafNodes]);

  useEffect(() => {
    if (!onSelectedItemIdsChange || !isCascadeSelection) {
      return;
    }

    // Keep parent IDs synchronized with descendant state in cascade mode.
    const normalizedSelectedItems = normalizeCascadeSelection(
      nodes,
      new Set(selectedItemIds),
      restrictSelectionToLeafNodes
    );

    if (!isSameSelection(selectedItemIds, normalizedSelectedItems)) {
      onSelectedItemIdsChange([...normalizedSelectedItems]);
    }
  }, [isCascadeSelection, nodes, onSelectedItemIdsChange, restrictSelectionToLeafNodes, selectedItemIds]);

  useEffect(() => {
    // Structural edits can leave stale branch IDs in openItems; prune them.
    const validBranchIds = new Set(collectBranchIds(nodes));
    setOpenItems((previousOpenItems) => {
      let changed = false;
      const nextOpenItems = new Set<TreeItemValue>();

      for (const item of previousOpenItems) {
        if (validBranchIds.has(item)) {
          nextOpenItems.add(item);
        } else {
          changed = true;
        }
      }

      return changed ? nextOpenItems : previousOpenItems;
    });
  }, [nodes]);

  useEffect(() => {
    if (editingId && !hasNode(nodes, editingId)) {
      setDrawerOpen(false);
      setEditingId(null);
      setPendingNewId(null);
      setDraftLabel("");
      setDraftValue("");
      setDraftTextHsvColor(DEFAULT_TEXT_HSV_COLOR);
      setDraftBackgroundHsvColor(DEFAULT_BACKGROUND_HSV_COLOR);
    }
  }, [nodes, editingId]);

  useEffect(() => {
    if (!isDragAndDropActive) {
      setDraggedNodeId(null);
      setDropHint(null);
    }
  }, [isDragAndDropActive]);

  useEffect(() => {
    if (draggedNodeId && !hasNode(nodes, draggedNodeId)) {
      setDraggedNodeId(null);
      setDropHint(null);
    }
  }, [draggedNodeId, nodes]);

  useEffect(() => {
    if (isSelectMode) {
      setActionsOverflowByNodeId({});
      return;
    }

    const rootElement = treeRootRef.current;
    if (!rootElement) {
      return;
    }

    const computeActionsOverflow = (): void => {
      const nextActionsOverflowByNodeId: Record<string, boolean> = {};
      const rowElements = rootElement.querySelectorAll<HTMLElement>("[data-node-row-id]");
      for (const rowElement of rowElements) {
        const rowNodeId = rowElement.dataset.nodeRowId;
        if (!rowNodeId) {
          continue;
        }

        nextActionsOverflowByNodeId[rowNodeId] = rowElement.clientWidth < FULL_ACTIONS_REQUIRED_WIDTH_PX;
      }

      setActionsOverflowByNodeId((previousActionsOverflowByNodeId) => {
        const previousKeys = Object.keys(previousActionsOverflowByNodeId);
        const nextKeys = Object.keys(nextActionsOverflowByNodeId);
        if (previousKeys.length === nextKeys.length) {
          let hasChanges = false;
          for (const nodeId of nextKeys) {
            if (previousActionsOverflowByNodeId[nodeId] !== nextActionsOverflowByNodeId[nodeId]) {
              hasChanges = true;
              break;
            }
          }

          if (!hasChanges) {
            return previousActionsOverflowByNodeId;
          }
        }

        return nextActionsOverflowByNodeId;
      });
    };

    computeActionsOverflow();

    if (typeof ResizeObserver === "undefined") {
      return;
    }

    const resizeObserver = new ResizeObserver(() => {
      computeActionsOverflow();
    });
    resizeObserver.observe(rootElement);

    return () => {
      resizeObserver.disconnect();
    };
  }, [isSelectMode, nodes, openItems, normalizedSearchQuery]);

  const closeEditor = (): void => {
    setDrawerOpen(false);
    setEditingId(null);
    setPendingNewId(null);
    setDraftLabel("");
    setDraftValue("");
    setDraftTextHsvColor(DEFAULT_TEXT_HSV_COLOR);
    setDraftBackgroundHsvColor(DEFAULT_BACKGROUND_HSV_COLOR);
  };

  const includeOpenItems = (values: Iterable<TreeItemValue>): void => {
    setOpenItems((previousOpenItems) => {
      const nextOpenItems = new Set(previousOpenItems);
      for (const value of values) {
        nextOpenItems.add(value);
      }
      return nextOpenItems;
    });
  };

  const openEditor = (node: TreeNode, isNew = false): void => {
    if (isSelectMode) {
      return;
    }

    setEditingId(node.id);
    setPendingNewId(isNew ? node.id : null);
    setDraftLabel(node.label);
    setDraftValue(node.value);
    setDraftTextHsvColor(hexToHsv(node.textColor));
    setDraftBackgroundHsvColor(hexToHsv(node.backgroundColor));
    setDrawerOpen(true);
  };

  const cancelEdit = (): void => {
    if (pendingNewId) {
      setNodes((prevNodes) => removeNodeById(prevNodes, pendingNewId));
    }

    closeEditor();
  };

  const saveEdit = (): void => {
    if (!editingId) {
      return;
    }

    const trimmed = draftLabel.trim();
    if (!trimmed) {
      return;
    }
    const nextValue = draftValue.trim();

    setNodes((prevNodes) =>
      updateNodeDetailsById(
        prevNodes,
        editingId,
        trimmed,
        nextValue,
        hsvToHex(draftTextHsvColor),
        hsvToHex(draftBackgroundHsvColor)
      )
    );
    closeEditor();
  };

  const addRootItem = (): void => {
    if (isSelectMode) {
      return;
    }

    const nextNode = createNode();
    setNodes((prevNodes) => [...prevNodes, nextNode]);
    openEditor(nextNode, true);
  };

  const addChildItem = (parentId: string): void => {
    if (isSelectMode) {
      return;
    }

    const nextNode = createNode();

    const parentPath = getPathById(nodes, parentId);
    // Ensure the destination branch is open so the new child is immediately visible.
    if (parentPath) {
      includeOpenItems(getPathNodeIds(nodes, parentPath));
    } else {
      includeOpenItems([parentId]);
    }

    setNodes((prevNodes) => addChildById(prevNodes, parentId, nextNode));
    openEditor(nextNode, true);
  };

  const editItem = (nodeId: string): void => {
    if (isSelectMode) {
      return;
    }

    const node = getNodeById(nodes, nodeId);
    if (!node) {
      return;
    }

    openEditor(node);
  };

  const moveItemUp = (path: NodePath): void => {
    setNodes((prevNodes) => moveNodeUpByPath(prevNodes, path));
  };

  const moveItemDown = (path: NodePath): void => {
    setNodes((prevNodes) => moveNodeDownByPath(prevNodes, path));
  };

  const indentItem = (path: NodePath): void => {
    const index = path[path.length - 1];
    if (index !== undefined && index > 0) {
      // Pre-open the new parent branch so the moved item stays visible.
      includeOpenItems(getPathNodeIds(nodes, [...path.slice(0, -1), index - 1]));
    }

    setNodes((prevNodes) => indentNodeByPath(prevNodes, path));
  };

  const outdentItem = (path: NodePath): void => {
    setNodes((prevNodes) => outdentNodeByPath(prevNodes, path));
  };

  const deleteItem = (nodeId: string): void => {
    setNodes((prevNodes) => removeNodeById(prevNodes, nodeId));

    if (editingId === nodeId || pendingNewId === nodeId) {
      closeEditor();
    }
  };

  const handleFormSubmit = (event: FormEvent<HTMLFormElement>): void => {
    event.preventDefault();
    saveEdit();
  };

  const handleTreeOpenChange = (_: unknown, data: TreeOpenChangeData): void => {
    setOpenItems(new Set(data.openItems));
  };

  const expandAll = (): void => {
    setOpenItems(new Set(branchNodeIds));
  };

  const collapseAll = (): void => {
    setOpenItems(new Set());
  };

  const toggleSelection = (node: TreeNode, checked: boolean, canSelectNode: boolean): void => {
    if (!onSelectedItemIdsChange) {
      return;
    }

    if (!canSelectNode) {
      return;
    }

    if (selectionMode === "single") {
      onSelectedItemIdsChange(checked ? [node.id] : []);
      return;
    }

    const nextSelectedItems = new Set(selectedItemsSet);
    if (isCascadeSelection) {
      // Cascade mode applies the toggle to the full selectable subtree.
      const relatedIds = collectSelectableSubtreeIds(node, restrictSelectionToLeafNodes, true);

      if (checked) {
        for (const relatedId of relatedIds) {
          nextSelectedItems.add(relatedId);
        }
      } else {
        for (const relatedId of relatedIds) {
          nextSelectedItems.delete(relatedId);
        }
      }
      onSelectedItemIdsChange([...normalizeCascadeSelection(nodes, nextSelectedItems, restrictSelectionToLeafNodes)]);
    } else if (checked) {
      nextSelectedItems.add(node.id);
      onSelectedItemIdsChange([...nextSelectedItems]);
    } else {
      nextSelectedItems.delete(node.id);
      onSelectedItemIdsChange([...nextSelectedItems]);
    }
  };

  const getDropPosition = (event: DragEvent<HTMLElement>): DropPosition => {
    const bounds = event.currentTarget.getBoundingClientRect();
    const yOffset = event.clientY - bounds.top;
    const edgeZone = Math.min(10, bounds.height * 0.3);

    if (yOffset < edgeZone) {
      return "before";
    }
    if (yOffset > bounds.height - edgeZone) {
      return "after";
    }
    return "inside";
  };

  const getDraggedNodeId = (event: DragEvent<HTMLElement>): string | null => {
    if (draggedNodeId) {
      return draggedNodeId;
    }

    const customTypeId = event.dataTransfer.getData(DRAG_NODE_MIME_TYPE).trim();
    if (customTypeId) {
      return customTypeId;
    }

    const plainTextId = event.dataTransfer.getData("text/plain").trim();
    return plainTextId || null;
  };

  const canDropOnTarget = (sourceId: string, targetId: string): boolean => {
    if (sourceId === targetId) {
      return false;
    }

    const sourcePath = getPathById(nodes, sourceId);
    const targetPath = getPathById(nodes, targetId);
    if (!sourcePath || !targetPath) {
      return false;
    }

    return !isPathPrefix(sourcePath, targetPath);
  };

  const handleDragStart = (event: DragEvent<HTMLElement>, nodeId: string): void => {
    if (!isDragAndDropActive) {
      return;
    }

    event.stopPropagation();
    event.dataTransfer.effectAllowed = "move";
    event.dataTransfer.setData(DRAG_NODE_MIME_TYPE, nodeId);
    event.dataTransfer.setData("text/plain", nodeId);
    setDraggedNodeId(nodeId);
    setDropHint(null);
  };

  const handleDragOver = (event: DragEvent<HTMLElement>, targetNode: TreeNode): void => {
    if (!isDragAndDropActive) {
      return;
    }

    const targetId = targetNode.id;
    const sourceId = getDraggedNodeId(event);
    if (!sourceId || !canDropOnTarget(sourceId, targetId)) {
      return;
    }

    event.preventDefault();
    event.stopPropagation();
    event.dataTransfer.dropEffect = "move";
    const position = getDropPosition(event);
    setDropHint((previousDropHint) => {
      if (previousDropHint?.targetId === targetId && previousDropHint.position === position) {
        return previousDropHint;
      }

      return { targetId, position };
    });
  };

  const handleDragLeave = (event: DragEvent<HTMLElement>, targetId: string): void => {
    const nextTarget = event.relatedTarget;
    if (nextTarget instanceof Node && event.currentTarget.contains(nextTarget)) {
      return;
    }

    setDropHint((previousDropHint) => (previousDropHint?.targetId === targetId ? null : previousDropHint));
  };

  const handleDrop = (event: DragEvent<HTMLElement>, targetNode: TreeNode): void => {
    if (!isDragAndDropActive) {
      return;
    }

    const targetId = targetNode.id;
    const sourceId = getDraggedNodeId(event);
    if (!sourceId) {
      setDropHint(null);
      setDraggedNodeId(null);
      return;
    }

    event.preventDefault();
    event.stopPropagation();

    if (!canDropOnTarget(sourceId, targetId)) {
      setDropHint(null);
      setDraggedNodeId(null);
      return;
    }

    const dropPosition = dropHint?.targetId === targetId ? dropHint.position : getDropPosition(event);
    if (dropPosition === "inside") {
      includeOpenItems([targetId]);
    }

    setNodes((prevNodes) => moveNodeByDrop(prevNodes, sourceId, targetId, dropPosition));
    setDropHint(null);
    setDraggedNodeId(null);
  };

  const handleDragEnd = (): void => {
    setDropHint(null);
    setDraggedNodeId(null);
  };

  const renderNode = (node: TreeNode, path: NodePath, siblingCount: number): JSX.Element | null => {
    // During search, render only matched nodes and their ancestors.
    if (isSearchActive && !searchVisibleNodeIds.has(node.id)) {
      return null;
    }

    const hasChildren = (node.children?.length ?? 0) > 0;
    const canSelectNode = !restrictSelectionToLeafNodes || !hasChildren;
    const isSearchMatch = isSearchActive && searchMatchingNodeIds.has(node.id);
    const checkboxState = isCascadeSelection ? (cascadeCheckStateById.get(node.id) ?? false) : selectedItemsSet.has(node.id);
    const index = path[path.length - 1] ?? 0;
    const canMoveUp = index > 0;
    const canMoveDown = index < siblingCount - 1;
    const canIndent = index > 0;
    const canOutdent = path.length > 1;
    const children = node.children ?? [];
    const dropPosition = dropHint?.targetId === node.id ? dropHint.position : null;
    const nodeRowClassName = `nodeRow${draggedNodeId === node.id ? " nodeRowDragging" : ""}${
      dropPosition === "before" ? " nodeRowDropBefore" : ""
    }${dropPosition === "inside" ? " nodeRowDropInside" : ""}${dropPosition === "after" ? " nodeRowDropAfter" : ""}`;
    const layoutClassName = `nodeTreeItemLayout${hasChildren ? " nodeTreeItemLayoutBranch" : " nodeTreeItemLayoutLeaf"}${
      isDragAndDropActive ? " nodeTreeItemLayoutWithDrag" : ""
    }`;

    return (
      <TreeItem key={node.id} value={node.id} itemType={hasChildren ? "branch" : "leaf"}>
        <TreeItemLayout
          className={layoutClassName}
          onDragOver={(event) => handleDragOver(event, node)}
          onDrop={(event) => handleDrop(event, node)}
          onDragLeave={(event) => handleDragLeave(event, node.id)}
          iconBefore={
            isDragAndDropActive
              ? {
                  className: "dragHandleSlot",
                  children: (
                    <button
                      type="button"
                      className="dragHandle"
                      draggable
                      aria-label={`Drag ${node.label}`}
                      tabIndex={-1}
                      onClick={(event) => event.stopPropagation()}
                      onDragStart={(event) => handleDragStart(event, node.id)}
                      onDragEnd={handleDragEnd}
                    >
                      <ReOrderDotsVerticalRegular className="dragHandleIcon" />
                    </button>
                  )
                }
              : undefined
          }
        >
          <div className={nodeRowClassName} data-node-row-id={node.id}>
            <div className="nodeLabelGroup">
              {isSelectMode ? (
                <Checkbox
                  className="selectionCheckbox"
                  checked={checkboxState}
                  disabled={!canSelectNode}
                  aria-label={`Select ${node.label}`}
                  onClick={(event) => event.stopPropagation()}
                  onChange={(_, data) => {
                    toggleSelection(node, data.checked === true, canSelectNode);
                  }}
                />
              ) : null}
              <button
                type="button"
                className={`nodePill nodePillButton${isSelectMode ? "" : " nodePillButtonEditable"}${
                  isSearchMatch ? " nodePillSearchMatch" : ""
                }`}
                style={{ color: node.textColor, backgroundColor: node.backgroundColor }}
                onClick={(event) => {
                  event.stopPropagation();
                  if (!isSelectMode) {
                    editItem(node.id);
                  }
                }}
              >
                {node.label}
              </button>
            </div>

            {!isSelectMode ? (
              <div className="actionGroup">
                {actionsOverflowByNodeId[node.id] ? (
                  <Menu positioning="below-end">
                    <MenuTrigger disableButtonEnhancement>
                      <Tooltip content="More actions" relationship="label" showDelay={TOOLTIP_SHOW_DELAY_MS}>
                        <Button
                          size="small"
                          appearance="subtle"
                          icon={<MoreHorizontalRegular />}
                          aria-label="More actions"
                          onClick={(event) => event.stopPropagation()}
                        />
                      </Tooltip>
                    </MenuTrigger>
                    <MenuPopover onClick={(event) => event.stopPropagation()}>
                      <MenuList>
                        <MenuItem
                          icon={<ArrowUpRegular />}
                          disabled={!canMoveUp}
                          onClick={(event) => {
                            event.stopPropagation();
                            moveItemUp(path);
                          }}
                        >
                          Move up
                        </MenuItem>
                        <MenuItem
                          icon={<ArrowDownRegular />}
                          disabled={!canMoveDown}
                          onClick={(event) => {
                            event.stopPropagation();
                            moveItemDown(path);
                          }}
                        >
                          Move down
                        </MenuItem>
                        <MenuItem
                          icon={<TextIndentIncreaseRegular />}
                          disabled={!canIndent}
                          onClick={(event) => {
                            event.stopPropagation();
                            indentItem(path);
                          }}
                        >
                          Indent
                        </MenuItem>
                        <MenuItem
                          icon={<TextIndentDecreaseRegular />}
                          disabled={!canOutdent}
                          onClick={(event) => {
                            event.stopPropagation();
                            outdentItem(path);
                          }}
                        >
                          Outdent
                        </MenuItem>
                      </MenuList>
                    </MenuPopover>
                  </Menu>
                ) : (
                  <>
                    <Tooltip content="Move up" relationship="label" showDelay={TOOLTIP_SHOW_DELAY_MS}>
                      <Button
                        size="small"
                        appearance="subtle"
                        icon={<ArrowUpRegular />}
                        disabled={!canMoveUp}
                        aria-label="Move up"
                        onClick={(event) => {
                          event.stopPropagation();
                          moveItemUp(path);
                        }}
                      />
                    </Tooltip>
                    <Tooltip content="Move down" relationship="label" showDelay={TOOLTIP_SHOW_DELAY_MS}>
                      <Button
                        size="small"
                        appearance="subtle"
                        icon={<ArrowDownRegular />}
                        disabled={!canMoveDown}
                        aria-label="Move down"
                        onClick={(event) => {
                          event.stopPropagation();
                          moveItemDown(path);
                        }}
                      />
                    </Tooltip>
                    <Tooltip content="Indent" relationship="label" showDelay={TOOLTIP_SHOW_DELAY_MS}>
                      <Button
                        size="small"
                        appearance="subtle"
                        icon={<TextIndentIncreaseRegular />}
                        disabled={!canIndent}
                        aria-label="Indent"
                        onClick={(event) => {
                          event.stopPropagation();
                          indentItem(path);
                        }}
                      />
                    </Tooltip>
                    <Tooltip content="Outdent" relationship="label" showDelay={TOOLTIP_SHOW_DELAY_MS}>
                      <Button
                        size="small"
                        appearance="subtle"
                        icon={<TextIndentDecreaseRegular />}
                        disabled={!canOutdent}
                        aria-label="Outdent"
                        onClick={(event) => {
                          event.stopPropagation();
                          outdentItem(path);
                        }}
                      />
                    </Tooltip>
                  </>
                )}
                <Tooltip content="Add child item" relationship="label" showDelay={TOOLTIP_SHOW_DELAY_MS}>
                  <Button
                    size="small"
                    appearance="subtle"
                    icon={<AddRegular />}
                    aria-label="Add child item"
                    onClick={(event) => {
                      event.stopPropagation();
                      addChildItem(node.id);
                    }}
                  />
                </Tooltip>
                <Tooltip content="Edit item" relationship="label" showDelay={TOOLTIP_SHOW_DELAY_MS}>
                  <Button
                    size="small"
                    appearance="subtle"
                    icon={<EditRegular />}
                    aria-label="Edit item"
                    onClick={(event) => {
                      event.stopPropagation();
                      editItem(node.id);
                    }}
                  />
                </Tooltip>
                <Tooltip content="Delete item" relationship="label" showDelay={TOOLTIP_SHOW_DELAY_MS}>
                  <Button
                    size="small"
                    appearance="subtle"
                    icon={<DeleteRegular />}
                    aria-label="Delete item"
                    onClick={(event) => {
                      event.stopPropagation();
                      deleteItem(node.id);
                    }}
                  />
                </Tooltip>
              </div>
            ) : null}
          </div>
        </TreeItemLayout>

        {hasChildren ? (
          <Tree aria-label={`${node.label} children`}>
            {children.map((child, childIndex) => renderNode(child, [...path, childIndex], children.length))}
          </Tree>
        ) : null}
      </TreeItem>
    );
  };

  return (
    <section className="editableTree" ref={treeRootRef}>
      <div className="toolbar">
        <Text weight="semibold">Items: {totalNodes}</Text>
        <div className="toolbarActions">
          <Tooltip content="Expand all" relationship="label" showDelay={TOOLTIP_SHOW_DELAY_MS}>
            <Button
              size="small"
              appearance="subtle"
              icon={<ArrowExpandAllRegular />}
              aria-label="Expand all"
              disabled={branchNodeIds.length === 0}
              onClick={expandAll}
            />
          </Tooltip>
          <Tooltip content="Collapse all" relationship="label" showDelay={TOOLTIP_SHOW_DELAY_MS}>
            <Button
              size="small"
              appearance="subtle"
              icon={<ArrowCollapseAllRegular />}
              aria-label="Collapse all"
              disabled={openItems.size === 0}
              onClick={collapseAll}
            />
          </Tooltip>
          {!isSelectMode ? (
            <Tooltip content="Add root item" relationship="label" showDelay={TOOLTIP_SHOW_DELAY_MS}>
              <Button
                size="small"
                appearance="primary"
                icon={<AddRegular />}
                aria-label="Add root item"
                onClick={addRootItem}
              />
            </Tooltip>
          ) : null}
        </div>
      </div>
      <div className="searchRow">
        <Input
          size="small"
          className="searchInput"
          value={searchQuery}
          contentBefore={<SearchRegular />}
          placeholder="Search tree items..."
          aria-label="Search tree items"
          onChange={(_, data) => setSearchQuery(data.value)}
        />
      </div>

      {nodes.length > 0 && (!isSearchActive || searchVisibleNodeIds.size > 0) ? (
        <Tree aria-label="Editable tree" openItems={effectiveOpenItems} onOpenChange={handleTreeOpenChange}>
          {nodes.map((node, index) => renderNode(node, [index], nodes.length))}
        </Tree>
      ) : (
        <Text className="emptyState">
          {isSearchActive ? `No items match "${searchQuery.trim()}".` : "Tree is empty. Add a root item to start."}
        </Text>
      )}

      <OverlayDrawer
        position="bottom"
        open={!isSelectMode && drawerOpen}
        modalType="modal"
        onOpenChange={(_, data) => {
          if (!data.open) {
            cancelEdit();
          }
        }}
      >
        <DrawerHeader>
          <DrawerHeaderTitle>{pendingNewId ? "Create tree item" : "Edit tree item"}</DrawerHeaderTitle>
        </DrawerHeader>
        <DrawerBody>
          <form id="tree-editor-form" className="editorForm" onSubmit={handleFormSubmit}>
            {editingPathNodes.length ? (
              <div className="pathBlock">
                <label className="fieldLabel">Path</label>
                <div className="pathPills">
                  {editingPathNodes.map((pathNode, pathIndex) => (
                    <span className="pathSegment" key={`${pathNode.id}-${pathIndex}`}>
                      <span className="nodePill pathPill" style={{ color: pathNode.textColor, backgroundColor: pathNode.backgroundColor }}>
                        {pathNode.label}
                      </span>
                      {pathIndex < editingPathNodes.length - 1 ? <span className="pathSeparator">&gt;</span> : null}
                    </span>
                  ))}
                </div>
              </div>
            ) : null}

            {editingId ? (
              <div className="pathBlock">
                <label className="fieldLabel">ID</label>
                <Text className="readonlyValue">{editingId}</Text>
              </div>
            ) : null}

            <label htmlFor="tree-item-label" className="fieldLabel">
              Label
            </label>
            <Input
              id="tree-item-label"
              value={draftLabel}
              autoFocus
              onChange={(_, data) => setDraftLabel(data.value)}
              aria-label="Tree item label"
            />

            <label htmlFor="tree-item-value" className="fieldLabel">
              Value
            </label>
            <Input
              id="tree-item-value"
              value={draftValue}
              onChange={(_, data) => setDraftValue(data.value)}
              aria-label="Tree item value"
            />

            <label className="fieldLabel">Text color</label>
            <div className="colorPickerBlock">
              <ColorPicker
                color={draftTextHsvColor}
                onColorChange={(_, data) => {
                  setDraftTextHsvColor({
                    h: data.color.h,
                    s: data.color.s,
                    v: data.color.v,
                    a: 1
                  });
                }}
              >
                <ColorArea />
                <ColorSlider />
              </ColorPicker>
              <div className="colorField">
                <span className="colorSwatch" style={{ backgroundColor: draftTextHexColor }} />
                <Text className="colorValue">{draftTextHexColor}</Text>
              </div>
            </div>

            <label className="fieldLabel">Background color</label>
            <div className="colorPickerBlock">
              <ColorPicker
                color={draftBackgroundHsvColor}
                onColorChange={(_, data) => {
                  setDraftBackgroundHsvColor({
                    h: data.color.h,
                    s: data.color.s,
                    v: data.color.v,
                    a: 1
                  });
                }}
              >
                <ColorArea />
                <ColorSlider />
              </ColorPicker>
              <div className="colorField">
                <span className="colorSwatch" style={{ backgroundColor: draftBackgroundHexColor }} />
                <Text className="colorValue">{draftBackgroundHexColor}</Text>
              </div>
            </div>

            <div className="pillPreviewWrap">
              <Text className="fieldLabel">Preview</Text>
              <span className="nodePill nodePillPreview" style={{ color: draftTextHexColor, backgroundColor: draftBackgroundHexColor }}>
                {draftLabel.trim() || "Tree item"}
              </span>
            </div>
          </form>
        </DrawerBody>
        <DrawerFooter>
          <Tooltip content="Cancel" relationship="label" showDelay={TOOLTIP_SHOW_DELAY_MS}>
            <Button appearance="secondary" icon={<DismissRegular />} aria-label="Cancel" onClick={cancelEdit} />
          </Tooltip>
          <Tooltip content="Save" relationship="label" showDelay={TOOLTIP_SHOW_DELAY_MS}>
            <Button
              appearance="primary"
              icon={<CheckmarkRegular />}
              aria-label="Save"
              type="submit"
              form="tree-editor-form"
              disabled={!draftLabel.trim()}
            />
          </Tooltip>
        </DrawerFooter>
      </OverlayDrawer>
    </section>
  );
}

export default EditableTree;
