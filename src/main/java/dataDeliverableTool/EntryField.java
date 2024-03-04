package dataDeliverableTool;

import java.awt.BorderLayout;
import java.awt.Color;
import java.awt.Dimension;
import java.awt.event.FocusEvent;
import java.awt.event.FocusListener;
import java.util.regex.Pattern;

import javax.swing.JLabel;
import javax.swing.JTextField;
import javax.swing.border.EmptyBorder;
import javax.swing.event.DocumentEvent;
import javax.swing.event.DocumentListener;
import javax.swing.text.Document;
import javax.swing.text.JTextComponent;

/**
 * A Field into which the user can input text, checked against a Regular Expression and containing "ghost" prompt text before the user focuses on it
 * @author Jaden
 *
 */
@SuppressWarnings("serial")
public class EntryField extends JTextField implements DocumentListener {
	/**
	 * The regular expression that this input field should match
	 */
	Pattern regex;
	/**
	 * Whether the text in this field is currently valid (i.e., matching {@link EntryField#regex})
	 */
	boolean isValid = false;
	
	/**
	 * Constructs an entry field with the specified regular expression and default (ghost) text
	 * @param regex the regular expression, as a String
	 * @param defaultText the ghost text
	 */
	EntryField(String regex, String defaultText) {
		super();
		TextPrompt prompt = new TextPrompt(defaultText, this);
		prompt.changeAlpha(150);
		this.setPreferredSize(new Dimension(prompt.getPreferredSize().width + 10, this.getPreferredSize().height));
		this.regex = Pattern.compile(regex);
		
		getDocument().addDocumentListener(this);
	}
	
	/**
	 * Checks if the text is currently valid, updating {@link EntryField#isValid}
	 */
	private void checkText() {
		String text = this.getText();
		if (regex.matcher(text).matches()) {
			this.setForeground(Color.BLACK);
			isValid = true;
		} else {
			this.setForeground(Color.RED);
			isValid = false;
		}
	}
	
	@Override
	public void changedUpdate(DocumentEvent arg0) {}

	@Override
	public void insertUpdate(DocumentEvent arg0) {
		checkText();
	}

	@Override
	public void removeUpdate(DocumentEvent arg0) {
		checkText();
	}
}

/**
 * "Ghost text" for an {@link EntryField} - disappears when the user is focusing on the text field
 * @author Jaden
 */
@SuppressWarnings("serial")
class TextPrompt extends JLabel implements FocusListener, DocumentListener {

	/**
	 * The EntryField that this is on
	 */
	private JTextComponent component;
	/**
	 * The Document of {@link TextPrompt#component}
	 */
	private Document document;

	/**
	 * Constructs a TextPrompt with the given text on the given EntryField
	 * @param text the text to show
	 * @param component the EntryField to be on
	 */
	public TextPrompt(String text, JTextComponent component) {
		this.component = component;
		document = component.getDocument();

		setText(text);
		setFont(component.getFont());
		setForeground(component.getForeground());
		setBorder(new EmptyBorder(component.getInsets()));
		setHorizontalAlignment(JLabel.LEADING);

		component.addFocusListener(this);
		document.addDocumentListener(this);

		component.setLayout(new BorderLayout());
		component.add(this);
		checkForPrompt();
	}

	/**
	 * Sets the alpha (transparency) of the TextPrompt
	 *
	 * @param alpha value in the range of 0 - 255.
	 */
	public void changeAlpha(int alpha) {
		alpha = alpha > 255 ? 255 : alpha < 0 ? 0 : alpha;

		Color foreground = getForeground();
		int red = foreground.getRed();
		int green = foreground.getGreen();
		int blue = foreground.getBlue();

		Color withAlpha = new Color(red, green, blue, alpha);
		super.setForeground(withAlpha);
	}

	/**
	 * Check whether the prompt should be visible or not
	 */
	private void checkForPrompt() {
		if (document.getLength() > 0) {
			setVisible(false);
			return;
		}

		if (component.hasFocus()) {
			setVisible(false);
		} else {
			setVisible(true);
		}
	}

	public void focusGained(FocusEvent e) {
		checkForPrompt();
	}

	public void focusLost(FocusEvent e) {
		checkForPrompt();
	}

	public void insertUpdate(DocumentEvent e) {
		checkForPrompt();
	}

	public void removeUpdate(DocumentEvent e) {
		checkForPrompt();
	}

	public void changedUpdate(DocumentEvent e) {
	}
}