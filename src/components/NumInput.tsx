import { useEffect, useState } from "react";

interface Props {
  value: number;
  onChange: (v: number) => void;
  /** Display multiplier, e.g. 100 to edit a fraction as percent. */
  scale?: number;
  min?: number;
  step?: number;
  className?: string;
  placeholder?: string;
  title?: string;
}

/** Number input that keeps the raw text while typing (allows "", "0.", "1,5"). */
export function NumInput({ value, onChange, scale = 1, min, step, className, placeholder, title }: Props) {
  const toText = (v: number) => (v ? String(+(v * scale).toFixed(6)) : "");
  const [text, setText] = useState(toText(value));
  const [focused, setFocused] = useState(false);

  useEffect(() => {
    if (!focused) setText(toText(value));
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [value, focused, scale]);

  return (
    <input
      className={`num ${className ?? ""}`}
      inputMode="decimal"
      value={text}
      placeholder={placeholder ?? "0"}
      title={title}
      onFocus={(e) => {
        setFocused(true);
        e.currentTarget.select();
      }}
      onBlur={() => setFocused(false)}
      onChange={(e) => {
        const t = e.target.value.replace(",", ".");
        setText(t);
        const v = t.trim() === "" ? 0 : Number(t);
        if (Number.isFinite(v)) {
          const scaled = v / scale;
          onChange(min !== undefined ? Math.max(min, scaled) : scaled);
        }
      }}
      step={step}
    />
  );
}
