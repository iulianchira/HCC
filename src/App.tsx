import { useState } from "react";
import { Button } from "@fluentui/react-button";
import { Text } from "@fluentui/react-text";
import EditableTree from "./EditableTree";
import type { TreeControlMode, TreeNode, TreeSelectionBehavior, TreeSelectionMode } from "./types";
import "./App.css";

const starterTree: TreeNode[] = [
  {
    id: "project",
    label: "Project",
    value: "project",
    textColor: "#FFFFFF",
    backgroundColor: "#0F6CBD",
    children: [
      { id: "requirements", label: "Requirements", value: "requirements", textColor: "#FFFFFF", backgroundColor: "#006C54" },
      {
        id: "implementation",
        label: "Implementation",
        value: "implementation",
        textColor: "#FFFFFF",
        backgroundColor: "#8A8886",
        children: [
          { id: "frontend", label: "Frontend", value: "frontend", textColor: "#FFFFFF", backgroundColor: "#5C2E91" },
          { id: "backend", label: "Backend", value: "backend", textColor: "#FFFFFF", backgroundColor: "#B146C2" }
        ]
      }
    ]
  },
  {
    id: "release",
    label: "Release",
    value: "release",
    textColor: "#FFFFFF",
    backgroundColor: "#CA5010",
    children: [{ id: "checklist", label: "Checklist", value: "checklist", textColor: "#FFFFFF", backgroundColor: "#986F0B" }]
  }
];

function App() {
  const [snapshot, setSnapshot] = useState<TreeNode[]>(starterTree);
  const [mode, setMode] = useState<TreeControlMode>("edit");
  const [selectionMode, setSelectionMode] = useState<TreeSelectionMode>("multiple");
  const [selectionBehavior, setSelectionBehavior] = useState<TreeSelectionBehavior>("independent");
  const [leafOnlySelection, setLeafOnlySelection] = useState(false);
  const [selectedItemIds, setSelectedItemIds] = useState<(string | number)[]>([]);

  const handleSelectionModeChange = (nextSelectionMode: TreeSelectionMode): void => {
    setSelectionMode(nextSelectionMode);
    if (nextSelectionMode === "single") {
      setSelectedItemIds((prevSelectedItemIds) => prevSelectedItemIds.slice(0, 1));
    }
  };

  return (
    <main className="appContainer">
      <header className="appHeader">
        <Text as="h1" size={700} weight="bold">
          Fluent UI v9 Editable Tree
        </Text>
        <Text>
          Switch between edit and select behavior. In select mode, checkbox selection supports independent or parent-child cascade behavior with optional leaf-only restriction.
        </Text>
      </header>

      <section className="controlPanel">
        <div className="controlRow">
          <Text weight="semibold">Behavior</Text>
          <div className="toggleGroup">
            <Button appearance={mode === "edit" ? "primary" : "secondary"} onClick={() => setMode("edit")}>
              Edit mode
            </Button>
            <Button appearance={mode === "select" ? "primary" : "secondary"} onClick={() => setMode("select")}>
              Select mode
            </Button>
          </div>
        </div>

        <div className="controlRow">
          <Text weight="semibold">Selection mode</Text>
          <div className="toggleGroup">
            <Button
              disabled={mode !== "select"}
              appearance={selectionMode === "single" ? "primary" : "secondary"}
              onClick={() => handleSelectionModeChange("single")}
            >
              Single
            </Button>
            <Button
              disabled={mode !== "select"}
              appearance={selectionMode === "multiple" ? "primary" : "secondary"}
              onClick={() => handleSelectionModeChange("multiple")}
            >
              Multiple
            </Button>
          </div>
          <Text className="selectionCount">Selected: {selectedItemIds.length}</Text>
        </div>

        <div className="controlRow">
          <Text weight="semibold">Parent-child behavior</Text>
          <div className="toggleGroup">
            <Button
              disabled={mode !== "select" || selectionMode !== "multiple"}
              appearance={selectionBehavior === "independent" ? "primary" : "secondary"}
              onClick={() => setSelectionBehavior("independent")}
            >
              Independent
            </Button>
            <Button
              disabled={mode !== "select" || selectionMode !== "multiple"}
              appearance={selectionBehavior === "cascade" ? "primary" : "secondary"}
              onClick={() => setSelectionBehavior("cascade")}
            >
              Cascade + mixed
            </Button>
          </div>
        </div>

        <div className="controlRow">
          <Text weight="semibold">Selectable items</Text>
          <div className="toggleGroup">
            <Button
              disabled={mode !== "select"}
              appearance={!leafOnlySelection ? "primary" : "secondary"}
              onClick={() => setLeafOnlySelection(false)}
            >
              Any node
            </Button>
            <Button
              disabled={mode !== "select"}
              appearance={leafOnlySelection ? "primary" : "secondary"}
              onClick={() => setLeafOnlySelection(true)}
            >
              Leaf only
            </Button>
          </div>
        </div>

      </section>

      <EditableTree
        initialNodes={starterTree}
        mode={mode}
        selectionMode={selectionMode}
        selectionBehavior={selectionBehavior}
        selectedItemIds={selectedItemIds}
        restrictSelectionToLeafNodes={leafOnlySelection}
        onSelectedItemIdsChange={setSelectedItemIds}
        onChange={setSnapshot}
      />

      <section className="snapshotBlock">
        <Text as="h2" size={500} weight="semibold">
          Current Tree JSON
        </Text>
        <pre>{JSON.stringify(snapshot, null, 2)}</pre>
      </section>

      <section className="snapshotBlock">
        <Text as="h2" size={500} weight="semibold">
          Selected Item IDs
        </Text>
        <pre>{JSON.stringify(selectedItemIds, null, 2)}</pre>
      </section>
    </main>
  );
}

export default App;
