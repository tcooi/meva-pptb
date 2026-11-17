import React from 'react';
import {
    Button,
    Input,
    Menu,
    MenuItemRadio,
    MenuGroupHeader,
    MenuList,
    MenuPopover,
    MenuTrigger,
    Spinner,
    makeStyles,
    mergeClasses,
    tokens,
} from '@fluentui/react-components';
import { BranchCompare16Regular, Box16Regular, ChevronDown12Regular, Filter16Regular, Search16Regular, Shapes16Regular, Warning16Regular } from '@fluentui/react-icons';
import type { SolutionInfo, VariableType } from '../services/environmentVariables';
import type { ColumnFilter } from '../services/columnFilters';
import { NO_TEXT_ASSIST } from './ui';

export type ValueFilter = 'all' | 'current' | 'default' | 'none';
export type CompareFilter = 'all' | 'different' | 'missing' | 'match';

export interface Filters {
    search: string;
    /** Per-column filters keyed by grid column id. */
    columnFilters: Record<string, ColumnFilter>;
    type: VariableType | 'all';
    solution: string;
    value: ValueFilter;
    compare: CompareFilter;
}

export const DEFAULT_FILTERS: Filters = { search: '', columnFilters: {}, type: 'all', solution: 'all', value: 'all', compare: 'all' };

const TYPES: VariableType[] = ['String', 'Number', 'Boolean', 'JSON', 'Data source', 'Secret'];

const VALUE_LABELS: Record<ValueFilter, string> = {
    all: 'Any value',
    current: 'Has current value',
    default: 'Using default value',
    none: 'No value set',
};

const COMPARE_LABELS: Record<CompareFilter, string> = {
    all: 'All variables',
    different: 'Different values',
    missing: 'Missing in an environment',
    match: 'Same everywhere',
};

const useStyles = makeStyles({
    root: {
        display: 'flex',
        flexWrap: 'wrap',
        alignItems: 'center',
        gap: '2px',
        minHeight: '44px',
        padding: `${tokens.spacingVerticalXS} ${tokens.spacingHorizontalM}`,
        borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
        backgroundColor: tokens.colorNeutralBackground1,
    },
    filterButton: {
        fontWeight: tokens.fontWeightRegular,
        color: tokens.colorNeutralForeground2,
        minWidth: 'auto',
    },
    // Tinted like a spreadsheet "Filtered by…" button so active filters are obvious
    active: {
        backgroundColor: tokens.colorPaletteLightGreenBackground1,
        color: tokens.colorNeutralForeground1,
        ':hover': {
            backgroundColor: tokens.colorPaletteLightGreenBackground2,
            color: tokens.colorNeutralForeground1,
        },
    },
    menuPopover: {
        maxHeight: '360px',
        overflowY: 'auto',
    },
    spacer: {
        flex: 1,
    },
    search: {
        width: '240px',
    },
    emptyOption: {
        color: tokens.colorNeutralForeground3,
    },
    count: {
        fontVariantNumeric: 'tabular-nums',
    },
    menuNote: {
        maxWidth: '320px',
        whiteSpace: 'normal',
        color: tokens.colorStatusWarningForeground3,
        fontWeight: tokens.fontWeightRegular,
    },
});

interface FilterMenuProps<T extends string> {
    name: string;
    label: string;
    icon: React.ReactElement;
    value: T;
    options: { value: T; label: string; secondary?: string; muted?: boolean }[];
    defaultValue: T;
    /** Shown above the options, e.g. to explain why the list is incomplete. */
    note?: string;
    disabled?: boolean;
    onChange: (value: T) => void;
}

/** A toolbar button that opens a single-choice menu; shows the chosen value and tints while active. */
function FilterMenu<T extends string>({ name, label, icon, value, options, defaultValue, disabled, note, onChange }: FilterMenuProps<T>) {
    const styles = useStyles();
    const isActive = value !== defaultValue;
    const selected = options.find((o) => o.value === value)?.label ?? value;

    return (
        <Menu checkedValues={{ [name]: [value] }} onCheckedValueChange={(_e, data) => onChange((data.checkedItems[0] ?? defaultValue) as T)}>
            <MenuTrigger disableButtonEnhancement>
                <Button
                    size="small"
                    appearance="subtle"
                    icon={icon}
                    disabled={disabled}
                    className={mergeClasses(styles.filterButton, isActive && styles.active)}
                    aria-label={`Filter by ${name}`}
                >
                    {isActive ? `${label}: ${selected}` : label}
                    <ChevronDown12Regular />
                </Button>
            </MenuTrigger>
            <MenuPopover className={styles.menuPopover}>
                <MenuList>
                    {note && <MenuGroupHeader className={styles.menuNote}>{note}</MenuGroupHeader>}
                    {options.map((option) => (
                        <MenuItemRadio
                            key={option.value}
                            name={name}
                            value={option.value}
                            className={option.muted ? styles.emptyOption : undefined}
                            secondaryContent={option.secondary && <span className={styles.count}>{option.secondary}</span>}
                        >
                            {option.label}
                        </MenuItemRadio>
                    ))}
                </MenuList>
            </MenuPopover>
        </Menu>
    );
}

interface FilterBarProps {
    filters: Filters;
    onChange: (filters: Filters) => void;
    /** All solutions, each with the number of environment variables it contains. */
    solutions: (SolutionInfo & { count: number })[];
    isLoadingSolutions: boolean;
    solutionsError?: string;
    isComparing: boolean;
}

export const FilterBar: React.FC<FilterBarProps> = ({ filters, onChange, solutions, isLoadingSolutions, solutionsError, isComparing }) => {
    const styles = useStyles();
    const set = <K extends keyof Filters>(key: K, value: Filters[K]) => onChange({ ...filters, [key]: value });
    const solutionsPending = isLoadingSolutions && solutions.length === 0;

    return (
        <div className={styles.root} role="toolbar" aria-label="Filters">
            <FilterMenu
                name="type"
                label="Type"
                icon={<Shapes16Regular />}
                value={filters.type}
                defaultValue="all"
                options={[{ value: 'all' as const, label: 'All types' }, ...TYPES.map((type) => ({ value: type, label: type }))]}
                onChange={(value) => set('type', value)}
            />
            <FilterMenu
                name="solution"
                label={solutionsPending ? 'Loading solutions…' : 'Solution'}
                icon={solutionsPending ? <Spinner size="extra-tiny" /> : solutionsError ? <Warning16Regular /> : <Box16Regular />}
                value={filters.solution}
                defaultValue="all"
                disabled={solutionsPending}
                note={solutionsError ? "Couldn't load solutions, so this list may be incomplete. See the notification for details." : undefined}
                options={[
                    { value: 'all', label: 'All solutions' },
                    ...solutions.map((s) => ({ value: s.uniqueName, label: s.friendlyName, secondary: String(s.count), muted: s.count === 0 })),
                ]}
                onChange={(value) => set('solution', value)}
            />
            <FilterMenu
                name="value"
                label="Value"
                icon={<Filter16Regular />}
                value={filters.value}
                defaultValue="all"
                options={Object.entries(VALUE_LABELS).map(([value, label]) => ({ value: value as ValueFilter, label }))}
                onChange={(value) => set('value', value)}
            />
            {isComparing && (
                <FilterMenu
                    name="comparison"
                    label="Comparison"
                    icon={<BranchCompare16Regular />}
                    value={filters.compare}
                    defaultValue="all"
                    options={Object.entries(COMPARE_LABELS).map(([value, label]) => ({ value: value as CompareFilter, label }))}
                    onChange={(value) => set('compare', value)}
                />
            )}

            <span className={styles.spacer} />

            <Input
                {...NO_TEXT_ASSIST}
                className={styles.search}
                size="small"
                appearance="filled-darker"
                contentBefore={<Search16Regular />}
                placeholder="Search all columns"
                value={filters.search}
                onChange={(_e, data) => set('search', data.value)}
                aria-label="Search environment variables"
            />
        </div>
    );
};
