export type TreeNode = {
  id: string;
  label: string;
  value: string;
  textColor: string;
  backgroundColor: string;
  children?: TreeNode[];
};

export type TreeControlMode = "edit" | "select";

export type TreeSelectionMode = "single" | "multiple";

export type TreeSelectionBehavior = "independent" | "cascade";
