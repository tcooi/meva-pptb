export type FilterOperator = 'contains' | 'notContains' | 'is' | 'isNot' | 'startsWith' | 'endsWith' | 'isEmpty' | 'isNotEmpty';

export interface ColumnFilter {
    operator: FilterOperator;
    value: string;
}

export const FILTER_OPERATORS: { value: FilterOperator; label: string; needsValue: boolean }[] = [
    { value: 'contains', label: 'contains', needsValue: true },
    { value: 'notContains', label: 'does not contain', needsValue: true },
    { value: 'is', label: 'is', needsValue: true },
    { value: 'isNot', label: 'is not', needsValue: true },
    { value: 'startsWith', label: 'starts with', needsValue: true },
    { value: 'endsWith', label: 'ends with', needsValue: true },
    { value: 'isEmpty', label: 'is empty', needsValue: false },
    { value: 'isNotEmpty', label: 'is not empty', needsValue: false },
];

/**
 * Stands in for a secret's value: it is never shown or matched as text,
 * but it still counts as "not empty".
 */
export const HIDDEN_VALUE = '\u0000hidden';

export function operatorNeedsValue(operator: FilterOperator): boolean {
    return FILTER_OPERATORS.find((o) => o.value === operator)?.needsValue ?? true;
}

/** A filter only applies once it has a value, unless its operator doesn't need one. */
export function isFilterActive(filter: ColumnFilter | undefined): filter is ColumnFilter {
    return !!filter && (!operatorNeedsValue(filter.operator) || filter.value.trim() !== '');
}

/**
 * Tests a column's lower-cased texts against a filter. A column can hold several texts
 * (e.g. an environment's current and default value): positive operators match if any text
 * matches, and the negative ones ("does not contain", "is not") match only if none does.
 */
export function matchesColumnFilter(texts: string[], filter: ColumnFilter): boolean {
    const needle = filter.value.trim().toLowerCase();
    const visible = texts.filter((t) => t !== HIDDEN_VALUE);
    switch (filter.operator) {
        case 'contains':
            return visible.some((t) => t.includes(needle));
        case 'notContains':
            return !visible.some((t) => t.includes(needle));
        case 'is':
            return visible.some((t) => t === needle);
        case 'isNot':
            return !visible.some((t) => t === needle);
        case 'startsWith':
            return visible.some((t) => t.startsWith(needle));
        case 'endsWith':
            return visible.some((t) => t.endsWith(needle));
        case 'isEmpty':
            return texts.every((t) => t === '');
        case 'isNotEmpty':
            return texts.some((t) => t !== '');
    }
}
