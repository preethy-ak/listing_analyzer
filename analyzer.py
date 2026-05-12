def analyze_data(df):
    summary = df.groupby("Issue Type").size().reset_index(name="Count")

    detailed = df.copy()

    return {
        "summary": summary,
        "detailed": detailed
    }
