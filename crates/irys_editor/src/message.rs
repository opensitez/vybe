use uuid::Uuid;
use irys_forms::ControlType;

#[derive(Debug, Clone)]
pub enum Message {
    // File menu
    NewProject,
    OpenProject,
    SaveProject,
    SaveProjectAs,
    CloseProject,
    Exit,

    // Edit menu
    Undo,
    Redo,
    Cut,
    Copy,
    Paste,
    Delete,
    SelectAll,
    Find,
    Replace,

    // View menu
    ToggleProjectExplorer,
    TogglePropertiesWindow,
    ToggleToolbox,
    ToggleImmediateWindow,

    // Project menu
    AddForm,
    AddModule,
    AddClass,
    RemoveForm(String),
    ProjectProperties,
    Components,

    // Run menu
    Start,
    Stop,
    Restart,
    StepInto,
    StepOver,

    // Window menu
    CascadeWindows,
    TileHorizontal,
    TileVertical,

    // Form management
    NewForm,
    SelectForm(String),
    GenerateEventHandlers,

    // Toolbox
    SelectTool(ControlType),

    // Designer
    CanvasClicked(i32, i32),
    ControlSelected(Uuid),
    ControlDoubleClicked(Uuid),
    FormSelected,
    ControlMoved(Uuid, i32, i32),
    ControlResized(Uuid, i32, i32),
    DesignerMouseDown(i32, i32),
    DesignerMouseMove(i32, i32),
    DesignerMouseUp,
    StartResize(Uuid, String), // control_id, handle position (e.g., "bottom-right")

    // Properties
    PropertyChanged(String, String),
    DeleteControl,

    // Code editor
    CodeChanged(String),
    ViewCode,
    ViewDesigner,

    // General
    None,
}
