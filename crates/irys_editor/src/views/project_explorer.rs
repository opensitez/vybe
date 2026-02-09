use crate::message::Message;
use crate::state::EditorState;
use iced::widget::{button, column, container, row, scrollable, text};
use iced::{Element, Length};

pub fn project_explorer(state: &EditorState) -> Element<'static, Message> {
    let mut col = column![
        text("Project Explorer").size(14),
    ]
    .spacing(5)
    .padding(10);

    if let Some(project) = &state.project {
        col = col.push(text(format!("📁 {}", project.name)).size(13));

        // Forms section
        col = col.push(text("").size(3));
        col = col.push(text("  📋 Forms").size(12));

        for form_module in &project.forms {
            let is_current = state.current_form.as_ref() == Some(&form_module.form.name);
            let prefix = if is_current { "  ▶ " } else { "    " };

            col = col.push(
                button(text(format!("{}{}", prefix, form_module.form.name)).size(11))
                    .on_press(Message::SelectForm(form_module.form.name.clone()))
                    .width(Length::Fill)
            );
        }

        // Modules section (placeholder)
        col = col.push(text("").size(3));
        col = col.push(text("  📝 Modules").size(12));

        for module in &project.modules {
            col = col.push(
                text(format!("    {}", module.name)).size(11)
            );
        }

        // Actions
        col = col.push(text("").size(5));
        col = col.push(
            button(text("Add Form").size(11))
                .on_press(Message::AddForm)
                .width(Length::Fill)
        );
        col = col.push(
            button(text("Add Module").size(11))
                .on_press(Message::AddModule)
                .width(Length::Fill)
        );
        col = col.push(
            button(text("Project Properties").size(11))
                .on_press(Message::ProjectProperties)
                .width(Length::Fill)
        );
    } else {
        col = col.push(text("No project loaded").size(12));
        col = col.push(text("").size(5));
        col = col.push(
            button(text("New Project").size(11))
                .on_press(Message::NewProject)
                .width(Length::Fill)
        );
    }

    container(
        scrollable(col)
    )
    .width(Length::Fixed(200.0))
    .height(Length::Fill)
    .into()
}
