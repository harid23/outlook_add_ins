function analyzeMail(body, recipients) {

    const fullPanRegex = /\b[A-Z]{5}[0-9]{4}[A-Z]\b/;
    const maskedPanRegex = /\b[A-Z]{5}\*{4}[0-9]{2}\b/;

    const hasFullPan = fullPanRegex.test(body);
    const hasMaskedPan = maskedPanRegex.test(body);

    const external = recipients.some(r =>
        !r.toLowerCase().endsWith("harid786@outlook.com")
    );

    if (hasFullPan && external) {
        return {
            action: "BLOCK",
            message: "Full PAN cannot be sent to external recipients."
        };
    }

    if (hasMaskedPan) {
        return {
            action: "WARN",
            message: "Masked PAN detected. Do you want to send anyway?"
        };
    }

    return { action: "ALLOW" };
}