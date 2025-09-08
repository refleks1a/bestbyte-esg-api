import matplotlib.pyplot as plt
import io


# --- Chart Generation Functions ---
def create_line_chart(df, x_col, y_col, title, ylabel):
    """Generates a line chart from a DataFrame."""
    plt.figure(figsize=(6, 4))
    plt.plot(df[x_col], df[y_col], marker='o')
    plt.title(title)
    plt.xlabel(x_col)
    plt.ylabel(ylabel)
    plt.grid(True)
    plt.xticks(rotation=45)
    plt.tight_layout()
    buf = io.BytesIO()
    plt.savefig(buf, format='png')
    plt.close()
    buf.seek(0)
    return buf


def create_bar_chart(df, x_col, y_cols, title):
    """Generates a bar chart from a DataFrame."""
    plt.figure(figsize=(6, 4))
    df.set_index(x_col)[y_cols].plot(kind='bar', ax=plt.gca())
    plt.title(title)
    plt.xlabel(x_col)
    plt.ylabel("Percentage (%)")
    plt.xticks(rotation=0)
    plt.legend(title="Type")
    plt.tight_layout()
    buf = io.BytesIO()
    plt.savefig(buf, format='png')
    plt.close()
    buf.seek(0)
    return buf


def create_pie_chart(labels, sizes, title):
    """Generates a pie chart."""
    plt.figure(figsize=(6, 4))
    plt.pie(sizes, labels=labels, autopct='%1.1f%%', startangle=90)
    plt.title(title)
    plt.axis('equal')  # Equal aspect ratio ensures that pie is drawn as a circle.
    plt.tight_layout()
    buf = io.BytesIO()
    plt.savefig(buf, format='png')
    plt.close()
    buf.seek(0)
    return buf
