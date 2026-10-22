"use strict";

const personSettings = {
  canonicalName(value) {
    const name = cleanText(value);
    const aliases = globalThis.HR_LOCAL_SETTINGS?.personAliases || {};
    return Object.hasOwn(aliases, name) ? cleanText(aliases[name]) || name : name;
  },

  confirmedHourly(person, month) {
    const canonical = this.canonicalName(person);
    const entry = (globalThis.HR_LOCAL_SETTINGS?.hourlyOverrides || []).find((row) =>
      this.canonicalName(row.person) === canonical && row.month === month);
    const hourly = Number(entry?.hourly);
    return Number.isFinite(hourly) && hourly > 0 ? hourly : 0;
  },
};
