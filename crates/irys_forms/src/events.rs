use serde::{Deserialize, Serialize};

#[derive(Debug, Clone, PartialEq, Eq, Hash, Serialize, Deserialize)]
pub enum EventType {
    Click,
    DblClick,
    DoubleClick,
    Load,
    Unload,
    Change,
    TextChanged,
    SelectedIndexChanged,
    CheckedChanged,
    ValueChanged,
    KeyPress,
    KeyDown,
    KeyUp,
    MouseClick,
    MouseDoubleClick,
    MouseDown,
    MouseUp,
    MouseMove,
    MouseEnter,
    MouseLeave,
    MouseWheel,
    GotFocus,
    LostFocus,
    Enter,
    Leave,
    Validated,
    Validating,
    Resize,
    Paint,
    FormClosing,
    FormClosed,
    Shown,
    Activated,
    Deactivate,
    Tick,
    Elapsed,
    Scroll,
    SelectedValueChanged,
    CellClick,
    CellDoubleClick,
    CellValueChanged,
    SelectionChanged,
    LinkClicked,
    SplitterMoved,
    SplitterMoving,
    DateChanged,
    DateSelected,
    ItemClicked,
    DropDownOpening,
    DropDownClosed,
    ColumnClick,
    NodeMouseClick,
    AfterSelect,
    BeforeSelect,
    ItemCheck,
    MaskInputRejected,
    // Additional common events
    DropDown,
    DropDownStyleChanged,
    DrawItem,
    MeasureItem,
    Format,
    DragDrop,
    DragEnter,
    DragLeave,
    DragOver,
    GiveFeedback,
    EnabledChanged,
    VisibleChanged,
    BackColorChanged,
    ForeColorChanged,
    FontChanged,
    SizeChanged,
    LocationChanged,
    TabIndexChanged,
    DockChanged,
    // DataGridView additional
    CellFormatting,
    CellPainting,
    CellContentClick,
    CellEndEdit,
    CellBeginEdit,
    CellValidating,
    CellEnter,
    CellLeave,
    DataError,
    RowEnter,
    RowLeave,
    RowValidating,
    RowValidated,
    ColumnHeaderMouseClick,
    RowHeaderMouseClick,
    CurrentCellChanged,
    DataBindingComplete,
    // TreeView additional
    NodeMouseDoubleClick,
    AfterCheck,
    BeforeCheck,
    AfterExpand,
    AfterCollapse,
    BeforeExpand,
    BeforeCollapse,
    AfterLabelEdit,
    BeforeLabelEdit,
    ItemDrag,
    // WebBrowser
    DocumentCompleted,
    Navigating,
    Navigated,
    ProgressChanged,
    // ListView
    ItemSelectionChanged,
    ItemActivate,
    ColumnWidthChanged,
    // TabControl
    SelectedIndexChanged2,
    Selected,
    Deselecting,
    Selecting,
    // Toolbar / StatusStrip
    ButtonClick,
    // Common
    HelpRequested,
    ContextMenuStripChanged,
    ParentChanged,
    HandleCreated,
    HandleDestroyed,
    Move,
    Disposed,
}

impl EventType {
    pub fn as_str(&self) -> &str {
        match self {
            EventType::Click => "Click",
            EventType::DblClick => "DblClick",
            EventType::DoubleClick => "DoubleClick",
            EventType::Load => "Load",
            EventType::Unload => "Unload",
            EventType::Change => "Change",
            EventType::TextChanged => "TextChanged",
            EventType::SelectedIndexChanged => "SelectedIndexChanged",
            EventType::CheckedChanged => "CheckedChanged",
            EventType::ValueChanged => "ValueChanged",
            EventType::KeyPress => "KeyPress",
            EventType::KeyDown => "KeyDown",
            EventType::KeyUp => "KeyUp",
            EventType::MouseClick => "MouseClick",
            EventType::MouseDoubleClick => "MouseDoubleClick",
            EventType::MouseDown => "MouseDown",
            EventType::MouseUp => "MouseUp",
            EventType::MouseMove => "MouseMove",
            EventType::MouseEnter => "MouseEnter",
            EventType::MouseLeave => "MouseLeave",
            EventType::MouseWheel => "MouseWheel",
            EventType::GotFocus => "GotFocus",
            EventType::LostFocus => "LostFocus",
            EventType::Enter => "Enter",
            EventType::Leave => "Leave",
            EventType::Validated => "Validated",
            EventType::Validating => "Validating",
            EventType::Resize => "Resize",
            EventType::Paint => "Paint",
            EventType::FormClosing => "FormClosing",
            EventType::FormClosed => "FormClosed",
            EventType::Shown => "Shown",
            EventType::Activated => "Activated",
            EventType::Deactivate => "Deactivate",
            EventType::Tick => "Tick",
            EventType::Elapsed => "Elapsed",
            EventType::Scroll => "Scroll",
            EventType::SelectedValueChanged => "SelectedValueChanged",
            EventType::CellClick => "CellClick",
            EventType::CellDoubleClick => "CellDoubleClick",
            EventType::CellValueChanged => "CellValueChanged",
            EventType::SelectionChanged => "SelectionChanged",
            EventType::LinkClicked => "LinkClicked",
            EventType::SplitterMoved => "SplitterMoved",
            EventType::SplitterMoving => "SplitterMoving",
            EventType::DateChanged => "DateChanged",
            EventType::DateSelected => "DateSelected",
            EventType::ItemClicked => "ItemClicked",
            EventType::DropDownOpening => "DropDownOpening",
            EventType::DropDownClosed => "DropDownClosed",
            EventType::ColumnClick => "ColumnClick",
            EventType::NodeMouseClick => "NodeMouseClick",
            EventType::AfterSelect => "AfterSelect",
            EventType::BeforeSelect => "BeforeSelect",
            EventType::ItemCheck => "ItemCheck",
            EventType::MaskInputRejected => "MaskInputRejected",
            EventType::DropDown => "DropDown",
            EventType::DropDownStyleChanged => "DropDownStyleChanged",
            EventType::DrawItem => "DrawItem",
            EventType::MeasureItem => "MeasureItem",
            EventType::Format => "Format",
            EventType::DragDrop => "DragDrop",
            EventType::DragEnter => "DragEnter",
            EventType::DragLeave => "DragLeave",
            EventType::DragOver => "DragOver",
            EventType::GiveFeedback => "GiveFeedback",
            EventType::EnabledChanged => "EnabledChanged",
            EventType::VisibleChanged => "VisibleChanged",
            EventType::BackColorChanged => "BackColorChanged",
            EventType::ForeColorChanged => "ForeColorChanged",
            EventType::FontChanged => "FontChanged",
            EventType::SizeChanged => "SizeChanged",
            EventType::LocationChanged => "LocationChanged",
            EventType::TabIndexChanged => "TabIndexChanged",
            EventType::DockChanged => "DockChanged",
            EventType::CellFormatting => "CellFormatting",
            EventType::CellPainting => "CellPainting",
            EventType::CellContentClick => "CellContentClick",
            EventType::CellEndEdit => "CellEndEdit",
            EventType::CellBeginEdit => "CellBeginEdit",
            EventType::CellValidating => "CellValidating",
            EventType::CellEnter => "CellEnter",
            EventType::CellLeave => "CellLeave",
            EventType::DataError => "DataError",
            EventType::RowEnter => "RowEnter",
            EventType::RowLeave => "RowLeave",
            EventType::RowValidating => "RowValidating",
            EventType::RowValidated => "RowValidated",
            EventType::ColumnHeaderMouseClick => "ColumnHeaderMouseClick",
            EventType::RowHeaderMouseClick => "RowHeaderMouseClick",
            EventType::CurrentCellChanged => "CurrentCellChanged",
            EventType::DataBindingComplete => "DataBindingComplete",
            EventType::NodeMouseDoubleClick => "NodeMouseDoubleClick",
            EventType::AfterCheck => "AfterCheck",
            EventType::BeforeCheck => "BeforeCheck",
            EventType::AfterExpand => "AfterExpand",
            EventType::AfterCollapse => "AfterCollapse",
            EventType::BeforeExpand => "BeforeExpand",
            EventType::BeforeCollapse => "BeforeCollapse",
            EventType::AfterLabelEdit => "AfterLabelEdit",
            EventType::BeforeLabelEdit => "BeforeLabelEdit",
            EventType::ItemDrag => "ItemDrag",
            EventType::DocumentCompleted => "DocumentCompleted",
            EventType::Navigating => "Navigating",
            EventType::Navigated => "Navigated",
            EventType::ProgressChanged => "ProgressChanged",
            EventType::ItemSelectionChanged => "ItemSelectionChanged",
            EventType::ItemActivate => "ItemActivate",
            EventType::ColumnWidthChanged => "ColumnWidthChanged",
            EventType::SelectedIndexChanged2 => "SelectedIndexChanged",
            EventType::Selected => "Selected",
            EventType::Deselecting => "Deselecting",
            EventType::Selecting => "Selecting",
            EventType::ButtonClick => "ButtonClick",
            EventType::HelpRequested => "HelpRequested",
            EventType::ContextMenuStripChanged => "ContextMenuStripChanged",
            EventType::ParentChanged => "ParentChanged",
            EventType::HandleCreated => "HandleCreated",
            EventType::HandleDestroyed => "HandleDestroyed",
            EventType::Move => "Move",
            EventType::Disposed => "Disposed",
        }
    }

    /// Return the .NET-compatible parameter signature for event handlers.
    pub fn parameters(&self) -> &'static str {
        match self {
            // Mouse events use MouseEventArgs
            EventType::MouseClick | EventType::MouseDoubleClick |
            EventType::MouseDown | EventType::MouseUp | EventType::MouseMove |
            EventType::MouseWheel =>
                "sender As Object, e As MouseEventArgs",
            // Key events
            EventType::KeyDown | EventType::KeyUp =>
                "sender As Object, e As KeyEventArgs",
            EventType::KeyPress =>
                "sender As Object, e As KeyPressEventArgs",
            // Form closing
            EventType::FormClosing =>
                "sender As Object, e As FormClosingEventArgs",
            EventType::FormClosed =>
                "sender As Object, e As FormClosedEventArgs",
            // Paint
            EventType::Paint =>
                "sender As Object, e As PaintEventArgs",
            // All other events use base EventArgs
            _ => "sender As Object, e As EventArgs",
        }
    }

    /// Parse an event name string into an EventType variant (case-insensitive).
    pub fn from_name(name: &str) -> Option<EventType> {
        match name.to_lowercase().as_str() {
            "click" => Some(EventType::Click),
            "dblclick" => Some(EventType::DblClick),
            "doubleclick" => Some(EventType::DoubleClick),
            "load" => Some(EventType::Load),
            "unload" => Some(EventType::Unload),
            "change" => Some(EventType::Change),
            "textchanged" => Some(EventType::TextChanged),
            "selectedindexchanged" => Some(EventType::SelectedIndexChanged),
            "checkedchanged" => Some(EventType::CheckedChanged),
            "valuechanged" => Some(EventType::ValueChanged),
            "keypress" => Some(EventType::KeyPress),
            "keydown" => Some(EventType::KeyDown),
            "keyup" => Some(EventType::KeyUp),
            "mouseclick" => Some(EventType::MouseClick),
            "mousedoubleclick" => Some(EventType::MouseDoubleClick),
            "mousedown" => Some(EventType::MouseDown),
            "mouseup" => Some(EventType::MouseUp),
            "mousemove" => Some(EventType::MouseMove),
            "mouseenter" => Some(EventType::MouseEnter),
            "mouseleave" => Some(EventType::MouseLeave),
            "mousewheel" => Some(EventType::MouseWheel),
            "gotfocus" => Some(EventType::GotFocus),
            "lostfocus" => Some(EventType::LostFocus),
            "enter" => Some(EventType::Enter),
            "leave" => Some(EventType::Leave),
            "validated" => Some(EventType::Validated),
            "validating" => Some(EventType::Validating),
            "resize" => Some(EventType::Resize),
            "paint" => Some(EventType::Paint),
            "formclosing" => Some(EventType::FormClosing),
            "formclosed" => Some(EventType::FormClosed),
            "shown" => Some(EventType::Shown),
            "activated" => Some(EventType::Activated),
            "deactivate" => Some(EventType::Deactivate),
            "tick" => Some(EventType::Tick),
            "elapsed" => Some(EventType::Elapsed),
            "scroll" => Some(EventType::Scroll),
            "selectedvaluechanged" => Some(EventType::SelectedValueChanged),
            "cellclick" => Some(EventType::CellClick),
            "celldoubleclick" => Some(EventType::CellDoubleClick),
            "cellvaluechanged" => Some(EventType::CellValueChanged),
            "selectionchanged" => Some(EventType::SelectionChanged),
            "linkclicked" => Some(EventType::LinkClicked),
            "splittermoved" => Some(EventType::SplitterMoved),
            "splittermoving" => Some(EventType::SplitterMoving),
            "datechanged" => Some(EventType::DateChanged),
            "dateselected" => Some(EventType::DateSelected),
            "itemclicked" => Some(EventType::ItemClicked),
            "dropdownopening" => Some(EventType::DropDownOpening),
            "dropdownclosed" => Some(EventType::DropDownClosed),
            "columnclick" => Some(EventType::ColumnClick),
            "nodemouseclick" => Some(EventType::NodeMouseClick),
            "afterselect" => Some(EventType::AfterSelect),
            "beforeselect" => Some(EventType::BeforeSelect),
            "itemcheck" => Some(EventType::ItemCheck),
            "maskinputrejected" => Some(EventType::MaskInputRejected),
            "dropdown" => Some(EventType::DropDown),
            "dropdownstylechanged" => Some(EventType::DropDownStyleChanged),
            "drawitem" => Some(EventType::DrawItem),
            "measureitem" => Some(EventType::MeasureItem),
            "format" => Some(EventType::Format),
            "dragdrop" => Some(EventType::DragDrop),
            "dragenter" => Some(EventType::DragEnter),
            "dragleave" => Some(EventType::DragLeave),
            "dragover" => Some(EventType::DragOver),
            "givefeedback" => Some(EventType::GiveFeedback),
            "enabledchanged" => Some(EventType::EnabledChanged),
            "visiblechanged" => Some(EventType::VisibleChanged),
            "backcolorchanged" => Some(EventType::BackColorChanged),
            "forecolorchanged" => Some(EventType::ForeColorChanged),
            "fontchanged" => Some(EventType::FontChanged),
            "sizechanged" => Some(EventType::SizeChanged),
            "locationchanged" => Some(EventType::LocationChanged),
            "tabindexchanged" => Some(EventType::TabIndexChanged),
            "dockchanged" => Some(EventType::DockChanged),
            "cellformatting" => Some(EventType::CellFormatting),
            "cellpainting" => Some(EventType::CellPainting),
            "cellcontentclick" => Some(EventType::CellContentClick),
            "cellendedit" => Some(EventType::CellEndEdit),
            "cellbeginedit" => Some(EventType::CellBeginEdit),
            "cellvalidating" => Some(EventType::CellValidating),
            "cellenter" => Some(EventType::CellEnter),
            "cellleave" => Some(EventType::CellLeave),
            "dataerror" => Some(EventType::DataError),
            "rowenter" => Some(EventType::RowEnter),
            "rowleave" => Some(EventType::RowLeave),
            "rowvalidating" => Some(EventType::RowValidating),
            "rowvalidated" => Some(EventType::RowValidated),
            "columnheadermouseclick" => Some(EventType::ColumnHeaderMouseClick),
            "rowheadermouseclick" => Some(EventType::RowHeaderMouseClick),
            "currentcellchanged" => Some(EventType::CurrentCellChanged),
            "databindingcomplete" => Some(EventType::DataBindingComplete),
            "nodemousedoubleclick" => Some(EventType::NodeMouseDoubleClick),
            "aftercheck" => Some(EventType::AfterCheck),
            "beforecheck" => Some(EventType::BeforeCheck),
            "afterexpand" => Some(EventType::AfterExpand),
            "aftercollapse" => Some(EventType::AfterCollapse),
            "beforeexpand" => Some(EventType::BeforeExpand),
            "beforecollapse" => Some(EventType::BeforeCollapse),
            "afterlabeledit" => Some(EventType::AfterLabelEdit),
            "beforelabeledit" => Some(EventType::BeforeLabelEdit),
            "itemdrag" => Some(EventType::ItemDrag),
            "documentcompleted" => Some(EventType::DocumentCompleted),
            "navigating" => Some(EventType::Navigating),
            "navigated" => Some(EventType::Navigated),
            "progresschanged" => Some(EventType::ProgressChanged),
            "itemselectionchanged" => Some(EventType::ItemSelectionChanged),
            "itemactivate" => Some(EventType::ItemActivate),
            "columnwidthchanged" => Some(EventType::ColumnWidthChanged),
            "selected" => Some(EventType::Selected),
            "deselecting" => Some(EventType::Deselecting),
            "selecting" => Some(EventType::Selecting),
            "buttonclick" => Some(EventType::ButtonClick),
            "helprequested" => Some(EventType::HelpRequested),
            "contextmenustripchanged" => Some(EventType::ContextMenuStripChanged),
            "parentchanged" => Some(EventType::ParentChanged),
            "handlecreated" => Some(EventType::HandleCreated),
            "handledestroyed" => Some(EventType::HandleDestroyed),
            "move" => Some(EventType::Move),
            "disposed" => Some(EventType::Disposed),
            _ => None,
        }
    }

    pub fn is_applicable_to(&self, control_type: Option<crate::ControlType>) -> bool {
        use crate::ControlType;
        
        match self {
            // Form-only events
            EventType::Load | EventType::Unload | EventType::FormClosing | EventType::FormClosed
            | EventType::Shown | EventType::Activated | EventType::Deactivate => control_type.is_none(),
            
            // Text change events
            EventType::Change | EventType::TextChanged => matches!(
                control_type,
                Some(ControlType::TextBox) | Some(ControlType::Label) | Some(ControlType::ComboBox)
                    | Some(ControlType::ListBox) | Some(ControlType::RichTextBox) | Some(ControlType::MaskedTextBox)
                    | None
            ),
            
            // Selection events (list-based)
            EventType::SelectedIndexChanged | EventType::SelectedValueChanged => matches!(
                control_type,
                Some(ControlType::ComboBox) | Some(ControlType::ListBox) | Some(ControlType::TabControl)
                    | Some(ControlType::ListView)
            ),

            // Checked state
            EventType::CheckedChanged => matches!(
                control_type,
                Some(ControlType::CheckBox) | Some(ControlType::RadioButton)
            ),

            // Value changed — broad: NumericUpDown, TrackBar, DateTimePicker, scrollbars, etc.
            EventType::ValueChanged => true,
            
            // Timer events
            EventType::Tick | EventType::Elapsed => true,

            // DataGridView cell events
            EventType::CellClick | EventType::CellDoubleClick | EventType::CellValueChanged
            | EventType::SelectionChanged | EventType::CellFormatting | EventType::CellPainting
            | EventType::CellContentClick | EventType::CellEndEdit | EventType::CellBeginEdit
            | EventType::CellValidating | EventType::CellEnter | EventType::CellLeave
            | EventType::DataError | EventType::RowEnter | EventType::RowLeave
            | EventType::RowValidating | EventType::RowValidated
            | EventType::ColumnHeaderMouseClick | EventType::RowHeaderMouseClick
            | EventType::CurrentCellChanged | EventType::DataBindingComplete => matches!(
                control_type,
                Some(ControlType::DataGridView)
            ),

            // LinkLabel
            EventType::LinkClicked => matches!(control_type, Some(ControlType::LinkLabel)),

            // SplitContainer
            EventType::SplitterMoved | EventType::SplitterMoving => matches!(control_type, Some(ControlType::SplitContainer)),

            // DateTimePicker / MonthCalendar
            EventType::DateChanged | EventType::DateSelected => matches!(
                control_type,
                Some(ControlType::DateTimePicker) | Some(ControlType::MonthCalendar)
            ),

            // ToolStrip / StatusStrip
            EventType::ItemClicked | EventType::ButtonClick => matches!(
                control_type,
                Some(ControlType::ToolStrip) | Some(ControlType::StatusStrip)
            ),

            // ComboBox dropdown events
            EventType::DropDown | EventType::DropDownStyleChanged => matches!(
                control_type,
                Some(ControlType::ComboBox)
            ),
            EventType::DropDownOpening | EventType::DropDownClosed => matches!(
                control_type,
                Some(ControlType::ComboBox) | Some(ControlType::ToolStripMenuItem)
            ),

            // Owner-draw events
            EventType::DrawItem | EventType::MeasureItem | EventType::Format => matches!(
                control_type,
                Some(ControlType::ComboBox) | Some(ControlType::ListBox) | Some(ControlType::ListView)
                    | Some(ControlType::TabControl)
            ),

            // ListView specific
            EventType::ColumnClick | EventType::ColumnWidthChanged
            | EventType::ItemSelectionChanged | EventType::ItemActivate => matches!(
                control_type,
                Some(ControlType::ListView)
            ),

            // TreeView specific
            EventType::NodeMouseClick | EventType::NodeMouseDoubleClick
            | EventType::AfterSelect | EventType::BeforeSelect
            | EventType::AfterCheck | EventType::BeforeCheck
            | EventType::AfterExpand | EventType::AfterCollapse
            | EventType::BeforeExpand | EventType::BeforeCollapse
            | EventType::AfterLabelEdit | EventType::BeforeLabelEdit
            | EventType::ItemDrag => matches!(
                control_type, Some(ControlType::TreeView)
            ),

            // ListView / CheckedListBox
            EventType::ItemCheck => matches!(
                control_type,
                Some(ControlType::ListView) | Some(ControlType::ListBox)
            ),

            // MaskedTextBox
            EventType::MaskInputRejected => matches!(control_type, Some(ControlType::MaskedTextBox)),

            // WebBrowser
            EventType::DocumentCompleted | EventType::Navigating | EventType::Navigated
            | EventType::ProgressChanged => matches!(control_type, Some(ControlType::WebBrowser)),

            // TabControl tab events
            EventType::SelectedIndexChanged2 | EventType::Selected
            | EventType::Deselecting | EventType::Selecting => matches!(
                control_type,
                Some(ControlType::TabControl)
            ),

            // Scroll — scrollbars, trackbars, panels, etc.
            EventType::Scroll => matches!(
                control_type,
                Some(ControlType::HScrollBar) | Some(ControlType::VScrollBar)
                    | Some(ControlType::TrackBar) | Some(ControlType::Panel)
                    | Some(ControlType::DataGridView) | Some(ControlType::RichTextBox)
                    | Some(ControlType::TextBox) | Some(ControlType::TreeView)
                    | Some(ControlType::ListView) | None
            ),

            // Click/double-click — universal
            EventType::Click | EventType::DblClick | EventType::DoubleClick
            | EventType::MouseClick | EventType::MouseDoubleClick => true,
            
            // Keyboard events — universal
            EventType::KeyPress | EventType::KeyDown | EventType::KeyUp => true,
            
            // Mouse events — universal
            EventType::MouseDown | EventType::MouseUp | EventType::MouseMove
            | EventType::MouseEnter | EventType::MouseLeave | EventType::MouseWheel => true,
            
            // Focus events — any control
            EventType::GotFocus | EventType::LostFocus | EventType::Enter | EventType::Leave
            | EventType::Validated | EventType::Validating => control_type.is_some(),

            // Drag events — universal
            EventType::DragDrop | EventType::DragEnter | EventType::DragLeave
            | EventType::DragOver | EventType::GiveFeedback => true,

            // Property changed events — universal
            EventType::EnabledChanged | EventType::VisibleChanged | EventType::BackColorChanged
            | EventType::ForeColorChanged | EventType::FontChanged | EventType::SizeChanged
            | EventType::LocationChanged | EventType::TabIndexChanged | EventType::DockChanged
            | EventType::ContextMenuStripChanged | EventType::ParentChanged => true,

            // Layout/paint/lifecycle — universal
            EventType::Resize | EventType::Paint | EventType::Move
            | EventType::HandleCreated | EventType::HandleDestroyed | EventType::Disposed
            | EventType::HelpRequested => true,
        }
    }

    pub fn all_events() -> Vec<EventType> {
        vec![
            EventType::Click,
            EventType::DblClick,
            EventType::DoubleClick,
            EventType::Load,
            EventType::Unload,
            EventType::Change,
            EventType::TextChanged,
            EventType::SelectedIndexChanged,
            EventType::CheckedChanged,
            EventType::ValueChanged,
            EventType::KeyPress,
            EventType::KeyDown,
            EventType::KeyUp,
            EventType::MouseClick,
            EventType::MouseDoubleClick,
            EventType::MouseDown,
            EventType::MouseUp,
            EventType::MouseMove,
            EventType::MouseEnter,
            EventType::MouseLeave,
            EventType::MouseWheel,
            EventType::GotFocus,
            EventType::LostFocus,
            EventType::Enter,
            EventType::Leave,
            EventType::Validated,
            EventType::Validating,
            EventType::Resize,
            EventType::Paint,
            EventType::FormClosing,
            EventType::FormClosed,
            EventType::Shown,
            EventType::Activated,
            EventType::Deactivate,
            EventType::Tick,
            EventType::Elapsed,
            EventType::Scroll,
            EventType::SelectedValueChanged,
            EventType::CellClick,
            EventType::CellDoubleClick,
            EventType::CellValueChanged,
            EventType::SelectionChanged,
            EventType::LinkClicked,
            EventType::SplitterMoved,
            EventType::SplitterMoving,
            EventType::DateChanged,
            EventType::DateSelected,
            EventType::ItemClicked,
            EventType::DropDownOpening,
            EventType::DropDownClosed,
            EventType::ColumnClick,
            EventType::NodeMouseClick,
            EventType::AfterSelect,
            EventType::BeforeSelect,
            EventType::ItemCheck,
            EventType::MaskInputRejected,
            EventType::DropDown,
            EventType::DropDownStyleChanged,
            EventType::DrawItem,
            EventType::MeasureItem,
            EventType::Format,
            EventType::DragDrop,
            EventType::DragEnter,
            EventType::DragLeave,
            EventType::DragOver,
            EventType::GiveFeedback,
            EventType::EnabledChanged,
            EventType::VisibleChanged,
            EventType::BackColorChanged,
            EventType::ForeColorChanged,
            EventType::FontChanged,
            EventType::SizeChanged,
            EventType::LocationChanged,
            EventType::TabIndexChanged,
            EventType::DockChanged,
            EventType::CellFormatting,
            EventType::CellPainting,
            EventType::CellContentClick,
            EventType::CellEndEdit,
            EventType::CellBeginEdit,
            EventType::CellValidating,
            EventType::CellEnter,
            EventType::CellLeave,
            EventType::DataError,
            EventType::RowEnter,
            EventType::RowLeave,
            EventType::RowValidating,
            EventType::RowValidated,
            EventType::ColumnHeaderMouseClick,
            EventType::RowHeaderMouseClick,
            EventType::CurrentCellChanged,
            EventType::DataBindingComplete,
            EventType::NodeMouseDoubleClick,
            EventType::AfterCheck,
            EventType::BeforeCheck,
            EventType::AfterExpand,
            EventType::AfterCollapse,
            EventType::BeforeExpand,
            EventType::BeforeCollapse,
            EventType::AfterLabelEdit,
            EventType::BeforeLabelEdit,
            EventType::ItemDrag,
            EventType::DocumentCompleted,
            EventType::Navigating,
            EventType::Navigated,
            EventType::ProgressChanged,
            EventType::ItemSelectionChanged,
            EventType::ItemActivate,
            EventType::ColumnWidthChanged,
            EventType::SelectedIndexChanged2,
            EventType::Selected,
            EventType::Deselecting,
            EventType::Selecting,
            EventType::ButtonClick,
            EventType::HelpRequested,
            EventType::ContextMenuStripChanged,
            EventType::ParentChanged,
            EventType::HandleCreated,
            EventType::HandleDestroyed,
            EventType::Move,
            EventType::Disposed,
        ]
    }
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct EventBinding {
    pub control_name: String,
    pub event_type: EventType,
    pub handler_name: String,
}

impl EventBinding {
    pub fn new(control_name: impl Into<String>, event_type: EventType) -> Self {
        let control_name = control_name.into();
        let handler_name = format!("{}_{}", control_name, event_type.as_str());

        Self {
            control_name,
            event_type,
            handler_name,
        }
    }

    pub fn with_handler(control_name: impl Into<String>, event_type: EventType, handler_name: impl Into<String>) -> Self {
        Self {
            control_name: control_name.into(),
            event_type,
            handler_name: handler_name.into(),
        }
    }
}
