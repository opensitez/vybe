use crate::message::Message;
use crate::state::EditorState;
use iced::widget::{button, column, container, row, text, text_input};
use iced::{Element, Length};

pub fn properties_panel(state: &EditorState) -> Element<'static, Message> {
    let mut col = column![text("Properties").size(16)]
        .spacing(10)
        .padding(10);

    // Show form properties if form is selected
    if state.form_selected {
        if let Some(form) = state.get_current_form() {
            col = col.push(text("Object: Form").size(14));
            col = col.push(text(format!("Name: {}", form.name)).size(13));

            col = col.push(text("Caption:").size(12));
            col = col.push(
                text_input("", &form.caption)
                    .on_input(|s| Message::PropertyChanged("FormCaption".to_string(), s))
                    .padding(5)
            );

            col = col.push(text("Size:").size(12));
            col = col.push(row![
                text("W:").size(11),
                text_input("", &form.width.to_string())
                    .on_input(|s| Message::PropertyChanged("FormWidth".to_string(), s))
                    .padding(3)
                    .width(Length::Fixed(60.0))
            ].spacing(5));

            col = col.push(row![
                text("H:").size(11),
                text_input("", &form.height.to_string())
                    .on_input(|s| Message::PropertyChanged("FormHeight".to_string(), s))
                    .padding(3)
                    .width(Length::Fixed(60.0))
            ].spacing(5));

            col = col.push(text("").size(5));
            col = col.push(text(format!("Controls: {}", form.controls.len())).size(12));
        }
    } else if let Some(control) = state.get_selected_control() {
        col = col.push(text(format!("Name: {}", control.name)).size(14));
        col = col.push(text(format!("Type: {}", control.control_type.as_str())).size(13));

        // Caption property
        if let Some(caption) = control.get_caption() {
            col = col.push(text("Caption:").size(12));
            col = col.push(
                text_input("", caption)
                    .on_input(|s| Message::PropertyChanged("Caption".to_string(), s))
                    .padding(5)
            );
        }

        // Text property
        if let Some(txt) = control.get_text() {
            col = col.push(text("Text:").size(12));
            col = col.push(
                text_input("", txt)
                    .on_input(|s| Message::PropertyChanged("Text".to_string(), s))
                    .padding(5)
            );
        }

        // Position and size
        col = col.push(text("Position & Size:").size(12));
        col = col.push(row![
            text("X:").size(11),
            text_input("", &control.bounds.x.to_string())
                .on_input(|s| Message::PropertyChanged("X".to_string(), s))
                .padding(3)
                .width(Length::Fixed(60.0))
        ].spacing(5));

        col = col.push(row![
            text("Y:").size(11),
            text_input("", &control.bounds.y.to_string())
                .on_input(|s| Message::PropertyChanged("Y".to_string(), s))
                .padding(3)
                .width(Length::Fixed(60.0))
        ].spacing(5));

        col = col.push(row![
            text("W:").size(11),
            text_input("", &control.bounds.width.to_string())
                .on_input(|s| Message::PropertyChanged("Width".to_string(), s))
                .padding(3)
                .width(Length::Fixed(60.0))
        ].spacing(5));

        col = col.push(row![
            text("H:").size(11),
            text_input("", &control.bounds.height.to_string())
                .on_input(|s| Message::PropertyChanged("Height".to_string(), s))
                .padding(3)
                .width(Length::Fixed(60.0))
        ].spacing(5));

        col = col.push(text("").size(10));
        col = col.push(
            button(text("Delete Control"))
                .on_press(Message::DeleteControl)
                .width(Length::Fill)
        );

    } else {
        col = col.push(text("No control selected").size(13));
    }

    container(col)
        .width(Length::Fixed(200.0))
        .height(Length::Fill)
        .into()
}
