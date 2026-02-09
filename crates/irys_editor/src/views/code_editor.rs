use crate::message::Message;
use crate::state::EditorState;
use iced::widget::{button, column, container, row, text, text_input};
use iced::{Element, Length};

pub fn code_editor_view(state: &EditorState) -> Element<'static, Message> {
    let code = state.get_current_code();

    let editor = column![
        row![
            button(text("Designer")).on_press(Message::ViewDesigner),
            text("Code Editor").size(16),
        ]
        .spacing(10),
        text_input("Enter irys code here...", &code)
            .on_input(Message::CodeChanged)
            .padding(10)
    ]
    .spacing(10)
    .padding(10);

    container(editor)
        .width(Length::Fill)
        .height(Length::Fill)
        .into()
}
