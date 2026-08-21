//! The control factory — one control kind name → one widget.
//!
//! A single place that turns a kind name into a live control, shared by every
//! surface that builds one: the `web:html` DOM (`<input type=checkbox>` →
//! `"checkbox"`), the .NET designer, and the Flutter realizer. Each of those
//! owns the mapping from ITS vocabulary to a kind name; none of them owns a
//! second copy of the construction.

use crate::layout::{LayoutRect, PanelWidget};
use crate::{
    BindingNavigator, Button, Checkbox, ContextMenu, DataGrid, DateTimePicker, FlowLayoutPanel,
    GroupBox, Label, ListBox, ListView, MaskedTextBox, MenuStrip, MonthCalendar, NumericUpDown,
    Panel, PictureBox, ProgressBar, Radio, ScrollBar, Select, Slider, SplitContainer, StatusStrip,
    TableLayoutPanel, Tabs, TextInput, ToolStrip, TreeView,
};

/// Create a boxed control from a kind name.
pub fn make_widget(
    type_name: &str,
    name: &str,
    text: &str,
    w: f32,
    h: f32,
) -> Box<dyn PanelWidget> {
    match type_name.to_lowercase().as_str() {
        "canvas" | "paintbox" => {
            // The Canvas widget is the bare drawable surface. PaintBox
            // is the .NET BCL/FCL alias the dotnet wrapper uses.

            let mut c = crate::Canvas::new().with_name(name);
            <crate::Canvas as PanelWidget>::set_rect(&mut c, LayoutRect::new(0.0, 0.0, w, h));
            Box::new(c)
        }
        "button" => {
            let mut b = Button::new(text).with_name(name);
            b.width = w;
            b.height = h;
            Box::new(b)
        }
        "label" | "linklabel" => {
            let mut l = Label::new(text).with_name(name);
            l.width = w;
            l.height = h;
            Box::new(l)
        }
        kind @ ("textbox" | "richtextbox") => {
            let mut t = TextInput::new().with_name(name);
            // A `<textarea>` / `TMemo` / `RichTextBox` is the same control with
            // newlines as content. `html/rich.rs` already declares it that way
            // (`white-space: pre-wrap`, `overflow: auto`); building it as a
            // one-line field was what made a memo show only its first line.
            t.multiline = kind == "richtextbox";
            t.value = text.to_string();
            t.cursor = t.value.len();
            t.width = w;
            t.height = h;
            Box::new(t)
        }
        "maskedtextbox" => {
            let mut t = MaskedTextBox::new().with_name(name);
            t.value = text.to_string();
            Box::new(t)
        }
        "checkbox" => Box::new(Checkbox::new(text).with_name(name)),
        "radiobutton" => Box::new(Radio::new(text).with_name(name)),
        "combobox" => {
            let mut s = Select::new(vec![]).with_name(name);
            s.width = w;
            s.height = h;
            Box::new(s)
        }
        "listbox" | "checkedlistbox" => {
            let mut l = ListBox::new().with_name(name);
            l.width = w;
            l.height = h;
            Box::new(l)
        }
        "panel" | "usercontrol" => {
            let mut p = Panel::new().with_name(name);
            p.width = w;
            p.height = h;
            Box::new(p)
        }
        "groupbox" | "frame" => {
            let mut g = GroupBox::new(text).with_name(name);
            g.width = w;
            g.height = h;
            Box::new(g)
        }
        "picturebox" => {
            let mut p = PictureBox::new().with_name(name);
            p.width = w;
            p.height = h;
            Box::new(p)
        }
        "progressbar" => {
            let mut p = ProgressBar::new().with_name(name);
            p.width = w;
            p.height = h;
            Box::new(p)
        }
        "trackbar" => Box::new(Slider::new(0.0, 100.0, 50.0).with_name(name)),
        "numericupdown" => Box::new(NumericUpDown::new().with_name(name)),
        "datetimepicker" => Box::new(DateTimePicker::new().with_name(name)),
        "colorpicker" => Box::new(crate::ColorPicker::new().with_name(name)),
        "fileinput" => Box::new(crate::FileInput::new().with_name(name)),
        "treeview" => Box::new(TreeView::new("", 1.0).with_name(name)),
        "datagridview" | "datagrid" => Box::new(DataGrid::new(&[]).with_name(name)),
        "listview" => Box::new(ListView::new().with_name(name)),
        "tabcontrol" => {
            let mut t = Tabs::new(&["Tab1"]).with_name(name);
            t.width = w;
            t.height = h;
            Box::new(t)
        }
        "monthcalendar" => Box::new(MonthCalendar::new().with_name(name)),
        "hscrollbar" => {
            let mut s = ScrollBar::new(false).with_name(name);
            s.width = w;
            s.height = h;
            Box::new(s)
        }
        "vscrollbar" => {
            let mut s = ScrollBar::new(true).with_name(name);
            s.width = w;
            s.height = h;
            Box::new(s)
        }
        "menustrip" => Box::new(MenuStrip::new().with_name(name)),
        "toolstrip" => Box::new(ToolStrip::new().with_name(name)),
        "statusstrip" => Box::new(StatusStrip::new().with_name(name)),
        "contextmenustrip" => Box::new(ContextMenu::new().with_name(name)),
        "splitcontainer" => Box::new(SplitContainer::new(false).with_name(name)),
        "flowlayoutpanel" => Box::new(FlowLayoutPanel::new().with_name(name)),
        // Horizontal flow — Flutter `Row`. (Vertical `FlowLayoutPanel` default
        // serves Column/Scaffold.)
        "hflowlayoutpanel" => Box::new(
            FlowLayoutPanel::new()
                .with_name(name)
                .with_direction(crate::flow_layout::FlowDirection::LeftToRight),
        ),
        "tablelayoutpanel" => Box::new(TableLayoutPanel::new(2, 2).with_name(name)),
        "bindingnavigator" => Box::new(BindingNavigator::new(name)),
        _ => {
            // Unknown control type — render as label placeholder
            let mut l = Label::new(&format!("[{}]", name)).with_name(name);
            l.transparent = true;
            l.width = w;
            l.height = h;
            Box::new(l)
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::layout::{CommandValue, WidgetCommand};

    /// Every text-bearing control must answer the font axis.
    ///
    /// A missing arm is invisible: `handle_command` returns `CommandValue::None`
    /// whether it acted or not, so a declaration reaching a widget with no arm
    /// is discarded in silence — which is what `Select`, `ListBox`, `Tabs` and
    /// `DataGrid` all did, `DataGrid` by having no `handle_command` at all.
    /// Reading the field back is the only way to tell the two apart.
    macro_rules! assert_font_reaches {
        ($widget:expr) => {{
            let mut w = $widget;
            w.handle_command(&WidgetCommand::Custom(
                "SetFontFamily".into(),
                CommandValue::Text("monospace".into()),
            ));
            w.handle_command(&WidgetCommand::Custom(
                "SetFontWeight".into(),
                CommandValue::Text("bold".into()),
            ));
            assert_eq!(
                (w.font.family.as_str(), w.font.weight),
                ("monospace", 700),
                concat!(stringify!($widget), " dropped a font declaration")
            );
        }};
    }

    #[test]
    fn a_font_declaration_reaches_every_text_bearing_control() {
        assert_font_reaches!(Button::new("ok"));
        assert_font_reaches!(Checkbox::new("on"));
        assert_font_reaches!(Label::new("hi"));
        assert_font_reaches!(Select::new(vec!["a".to_string()]));
        assert_font_reaches!(ListBox::new());
        assert_font_reaches!(Tabs::new(&["one"]));
        assert_font_reaches!(DataGrid::new(&["col"]));
        assert_font_reaches!(TextInput::new());
    }
}
