use crate::message::Message;
use crate::state::EditorState;
use iced::widget::{button, column, container, row, scrollable, text, text_editor};
use iced::{Element, Length};

pub fn code_editor_view(state: &EditorState) -> Element<'static, Message> {
    let code = state.get_current_code();

    // Create a text editor content
    let content = text_editor::Content::with_text(&code);

    let editor_widget = column![
        row![
            button(text("Designer").size(13))
                .on_press(Message::ViewDesigner)
                .padding(8),
            text("Code Editor").size(16),
            text(format!("Form: {}", state.current_form.as_ref().unwrap_or(&"None".to_string())))
                .size(13)
                .style(|_theme| {
                    text::Style {
                        color: Some(iced::Color::from_rgb(0.5, 0.5, 0.5)),
                    }
                }),
        ]
        .spacing(10)
        .padding(10),
        container(
            scrollable(
                text_editor(&content)
                    .on_action(|action| {
                        // Handle text editor actions
                        Message::None
                    })
                    .padding(10)
            )
        )
        .width(Length::Fill)
        .height(Length::Fill)
        .style(|_theme| {
            container::Style {
                background: Some(iced::Background::Color(iced::Color::WHITE)),
                border: iced::Border {
                    color: iced::Color::from_rgb(0.8, 0.8, 0.8),
                    width: 1.0,
                    ..Default::default()
                },
                ..Default::default()
            }
        }),
        row![
            text("Note: Full text editing coming soon. For now, use the simple demo code.")
                .size(11)
                .style(|_theme| {
                    text::Style {
                        color: Some(iced::Color::from_rgb(0.6, 0.6, 0.6)),
                    }
                }),
        ]
        .padding(10),
    ]
    .spacing(5);

    container(editor_widget)
        .width(Length::Fill)
        .height(Length::Fill)
        .padding(10)
        .into()
}
