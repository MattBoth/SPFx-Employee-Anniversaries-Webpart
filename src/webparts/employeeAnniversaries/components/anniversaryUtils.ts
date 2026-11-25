const MS_PER_DAY = 86400000;

export interface IAnniversaryMatch {
  date: Date;
  offsetDays: number; // negative = past, 0 = today, positive = upcoming
  years: number;
}

function startOfDay(d: Date): Date {
  return new Date(d.getFullYear(), d.getMonth(), d.getDate());
}

function daysBetween(from: Date, to: Date): number {
  const a = Date.UTC(from.getFullYear(), from.getMonth(), from.getDate());
  const b = Date.UTC(to.getFullYear(), to.getMonth(), to.getDate());
  return Math.round((b - a) / MS_PER_DAY);
}

// Anniversary in a given year; 29 Feb falls back to 28 Feb in non-leap years.
function anniversaryInYear(start: Date, year: number): Date {
  const month = start.getMonth();
  const day = start.getDate();
  const lastDay = new Date(year, month + 1, 0).getDate();
  return new Date(year, month, Math.min(day, lastDay));
}

export function findAnniversary(
  start: Date,
  today: Date,
  daysBefore: number,
  daysAfter: number
): IAnniversaryMatch | undefined {
  const t = startOfDay(today);
  // Check neighbouring years so windows spanning New Year are handled.
  for (const year of [t.getFullYear() - 1, t.getFullYear(), t.getFullYear() + 1]) {
    const years = year - start.getFullYear();
    if (years < 1) {
      continue;
    }
    const date = anniversaryInYear(start, year);
    const offsetDays = daysBetween(t, date);
    if (offsetDays >= -daysBefore && offsetDays <= daysAfter) {
      return { date, offsetDays, years };
    }
  }
  return undefined;
}
