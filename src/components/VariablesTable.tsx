import React, { useCallback, useRef } from 'react';
import { Button, Input, Popover, PopoverSurface, PopoverTrigger, Select, Table, TableBody, TableHeader, TableHeaderCell, TableRow, Text, Tooltip, makeStyles, mergeClasses, tokens } from '@fluentui/react-components';
import { BranchCompare16Regular, Code16Regular, Filter16Filled, Filter16Regular, Shapes16Regular, TextT16Regular } from '@fluentui/react-icons';
import { FILTER_OPERATORS, isFilterActive, operatorNeedsValue, type ColumnFilter, type FilterOperator } from '../services/columnFilters';
import { canCopy, type EditTarget, type EnvironmentSnapshot, type EnvironmentVariable, type ValueKind, type VariableRow } from '../services/environmentVariables';
import { EnvDot, NO_TEXT_ASSIST, StatusPill, TypePill } from './ui';
import { MIN_COLUMN_WIDTH, useColumnWidths } from '../hooks/useColumnWidths';

// Fluent's Edit16Regular / Edit16Filled paths, drawn as a CSS mask on the grid's edit buttons so each
// pencil is a single element instead of two icon components (the grid can hold thousands of them)
const pencilMask = (path: string) =>
    `url("data:image/svg+xml,${encodeURIComponent(`<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 16 16"><path d="${path}"/></svg>`)}")`;
const PENCIL_REGULAR = pencilMask(
    'M14.24 1.76a2.6 2.6 0 0 0-3.71 0L2.66 9.64q-.57.57-.78 1.35l-.86 3.39a.5.5 0 0 0 .6.6l3.39-.86q.78-.21 1.35-.78l7.88-7.87a2.6 2.6 0 0 0 0-3.7m-3 .71a1.62 1.62 0 1 1 2.29 2.3l-.78.77-2.3-2.29zM9.75 3.96l2.3 2.29-6.4 6.39q-.37.37-.89.5l-2.57.67.66-2.57q.14-.52.51-.9z'
);
const PENCIL_FILLED = pencilMask(
    'M10.53 1.76a2.62 2.62 0 1 1 3.7 3.71l-.77.78-3.71-3.7zM9.04 3.25 2.66 9.64q-.57.57-.78 1.35l-.86 3.39a.5.5 0 0 0 .6.6l3.39-.86q.78-.21 1.35-.78l6.39-6.38z'
);

const useStyles = makeStyles({
    container: {
        flex: 1,
        overflow: 'auto',
        backgroundColor: tokens.colorNeutralBackground1,
    },
    table: {
        tableLayout: 'fixed',
        borderCollapse: 'separate',
        borderSpacing: 0,
    },
    header: {
        position: 'sticky',
        top: 0,
        zIndex: 1,
    },
    // Spreadsheet-style header: light grey, normal case, field icon before the name
    headerCell: {
        height: '32px',
        padding: `0 ${tokens.spacingHorizontalS}`,
        backgroundColor: tokens.colorNeutralBackground2,
        borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
        borderRight: `1px solid ${tokens.colorNeutralStroke2}`,
        fontSize: tokens.fontSizeBase200,
        fontWeight: tokens.fontWeightSemibold,
        color: tokens.colorNeutralForeground2,
    },
    resizableHeader: {
        position: 'relative',
        // Fluent's inner wrapper is position: relative too, which would anchor the resize handle at the
        // padding edge; making it static anchors the handle to the cell itself so it straddles the border
        '& > .fui-TableHeaderCell__button': {
            position: 'static',
        },
    },
    // Invisible grab area centred on the column's right border; shows a line on hover/drag/focus
    resizeHandle: {
        position: 'absolute',
        top: 0,
        right: '-5px',
        bottom: 0,
        width: '9px',
        zIndex: 2,
        cursor: 'col-resize',
        touchAction: 'none',
        outline: 'none',
        '::after': {
            content: '""',
            position: 'absolute',
            top: '4px',
            bottom: '4px',
            left: '3px',
            width: '3px',
            borderRadius: '1px',
            backgroundColor: 'transparent',
        },
        ':hover::after': {
            backgroundColor: tokens.colorBrandStroke1,
        },
        ':focus-visible::after': {
            backgroundColor: tokens.colorBrandStroke1,
        },
    },
    resizing: {
        '::after': {
            backgroundColor: tokens.colorBrandStroke1,
        },
    },
    // Stretches across the cell so the filter button sits at the right edge
    headerInner: {
        display: 'flex',
        alignItems: 'center',
        gap: '2px',
        flex: 1,
        width: '100%',
        minWidth: 0,
    },
    headerLabel: {
        flex: 1,
        minWidth: 0,
    },
    filterButton: {
        flexShrink: 0,
        minWidth: '22px',
        width: '22px',
        height: '22px',
        color: tokens.colorNeutralForeground3,
    },
    filterButtonActive: {
        color: tokens.colorBrandForeground1,
        backgroundColor: tokens.colorBrandBackground2,
        ':hover': {
            color: tokens.colorBrandForeground1,
            backgroundColor: tokens.colorBrandBackground2Hover,
        },
    },
    filterSurface: {
        display: 'flex',
        flexDirection: 'column',
        gap: tokens.spacingVerticalS,
        width: '260px',
        padding: tokens.spacingHorizontalM,
    },
    filterActions: {
        display: 'flex',
        justifyContent: 'flex-end',
        gap: tokens.spacingHorizontalS,
    },
    emptyState: {
        display: 'flex',
        minHeight: '240px',
    },
    filler: {
        padding: 0,
        borderBottom: `1px solid ${tokens.colorNeutralStroke3}`,
    },
    fillerHeader: {
        padding: 0,
        backgroundColor: tokens.colorNeutralBackground2,
        borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
    },
    headerContent: {
        display: 'flex',
        alignItems: 'center',
        gap: '6px',
        minWidth: 0,
        '& svg': {
            flexShrink: 0,
            color: tokens.colorNeutralForeground3,
        },
    },
    row: {
        cursor: 'pointer',
        color: tokens.colorNeutralForeground1,
        ':hover': {
            backgroundColor: tokens.colorNeutralBackground1Hover,
        },
        ':focus-visible': {
            outline: `2px solid ${tokens.colorStrokeFocus2}`,
            outlineOffset: '-2px',
        },
    },
    cell: {
        height: '34px',
        padding: `0 ${tokens.spacingHorizontalS}`,
        borderBottom: `1px solid ${tokens.colorNeutralStroke3}`,
        borderRight: `1px solid ${tokens.colorNeutralStroke3}`,
        fontSize: tokens.fontSizeBase300,
        verticalAlign: 'middle',
        overflow: 'hidden',
        textOverflow: 'ellipsis',
        whiteSpace: 'nowrap',
    },
    tallCell: {
        height: '50px',
    },
    name: {
        fontWeight: tokens.fontWeightSemibold,
    },
    secondary: {
        color: tokens.colorNeutralForeground3,
    },
    primary: {
        color: tokens.colorNeutralForeground1,
    },
    // Translucent tints stay subtle on both light and dark surfaces
    differentCell: {
        backgroundColor: `color-mix(in srgb, ${tokens.colorStatusWarningBackground3} 9%, transparent)`,
    },
    missingCell: {
        backgroundColor: `color-mix(in srgb, ${tokens.colorStatusDangerBackground3} 9%, transparent)`,
    },
    // Pencil (or Copy here) appears when the cell is hovered or the button is focused
    valueCell: {
        ':hover [data-edit-button]': {
            opacity: 1,
        },
        ':hover [data-copy-button]': {
            opacity: 1,
        },
    },
    // Same lightweight native button as the pencil, with a text label
    copyButton: {
        flexShrink: 0,
        height: '22px',
        padding: `0 ${tokens.spacingHorizontalS}`,
        border: 'none',
        borderRadius: tokens.borderRadiusMedium,
        backgroundColor: 'transparent',
        color: tokens.colorBrandForeground1,
        fontFamily: tokens.fontFamilyBase,
        fontSize: tokens.fontSizeBase200,
        fontWeight: tokens.fontWeightSemibold,
        cursor: 'pointer',
        opacity: 0,
        ':hover': {
            backgroundColor: tokens.colorBrandBackground2,
        },
        ':hover:active': {
            backgroundColor: tokens.colorBrandBackground2Pressed,
        },
        ':focus-visible': {
            opacity: 1,
            outline: `2px solid ${tokens.colorStrokeFocus2}`,
            outlineOffset: '-2px',
        },
    },
    editLine: {
        display: 'flex',
        alignItems: 'center',
        gap: '2px',
        minWidth: 0,
    },
    editLineText: {
        flex: 1,
        minWidth: 0,
        overflow: 'hidden',
        textOverflow: 'ellipsis',
        whiteSpace: 'nowrap',
    },
    // A native button rather than Fluent's Button + Tooltip: the grid has one per value line (thousands
    // with several environments), and the Fluent components made filtering/re-rendering visibly slow.
    editButton: {
        position: 'relative',
        flexShrink: 0,
        display: 'inline-flex',
        alignItems: 'center',
        justifyContent: 'center',
        width: '22px',
        height: '22px',
        padding: 0,
        border: 'none',
        borderRadius: tokens.borderRadiusMedium,
        backgroundColor: 'transparent',
        color: tokens.colorNeutralForeground3,
        cursor: 'pointer',
        opacity: 0,
        // The pencil itself: outline normally, filled while hovered
        '::before': {
            content: '""',
            width: '16px',
            height: '16px',
            backgroundColor: 'currentColor',
            maskImage: PENCIL_REGULAR,
            maskRepeat: 'no-repeat',
            maskSize: '16px 16px',
        },
        // Clearly highlighted while the pencil itself is hovered, not just the cell
        ':hover': {
            color: tokens.colorBrandForeground1,
            backgroundColor: tokens.colorBrandBackground2,
        },
        ':hover::before': {
            maskImage: PENCIL_FILLED,
        },
        ':hover:active': {
            backgroundColor: tokens.colorBrandBackground2Pressed,
        },
        ':focus-visible': {
            opacity: 1,
            outline: `2px solid ${tokens.colorStrokeFocus2}`,
            outlineOffset: '-2px',
        },
        // Tooltip in Fluent's style, drawn with CSS so it costs nothing until shown. It sits to the left of the
        // pencil, inside the cell, so it isn't cut off under the sticky header on the first rows.
        '::after': {
            content: 'attr(data-tooltip)',
            position: 'absolute',
            right: 'calc(100% + 8px)',
            top: '50%',
            transform: 'translateY(-50%)',
            zIndex: 3,
            padding: '4px 11px 6px',
            borderRadius: tokens.borderRadiusMedium,
            backgroundColor: tokens.colorNeutralBackground1,
            color: tokens.colorNeutralForeground1,
            boxShadow: tokens.shadow16,
            fontFamily: tokens.fontFamilyBase,
            fontSize: tokens.fontSizeBase200,
            fontWeight: tokens.fontWeightRegular,
            lineHeight: tokens.lineHeightBase200,
            whiteSpace: 'nowrap',
            pointerEvents: 'none',
            opacity: 0,
            visibility: 'hidden',
            transitionProperty: 'opacity, visibility',
            transitionDuration: tokens.durationFaster,
        },
        ':hover::after': {
            opacity: 1,
            visibility: 'visible',
            transitionDelay: '250ms',
        },
        ':focus-visible::after': {
            opacity: 1,
            visibility: 'visible',
        },
    },
    lines: {
        display: 'flex',
        flexDirection: 'column',
        gap: '1px',
        minWidth: 0,
    },
    truncate: {
        overflow: 'hidden',
        textOverflow: 'ellipsis',
        whiteSpace: 'nowrap',
    },
    small: {
        fontSize: tokens.fontSizeBase200,
    },
    empty: {
        color: tokens.colorNeutralForeground4,
    },
});

type Styles = ReturnType<typeof useStyles>;


/** Secrets hold a Key Vault reference, not the secret itself, but we still avoid showing it in the grid. */
function displayValue(variable: EnvironmentVariable, value: string | null): string | null {
    if (value === null || value === '') {
        return value;
    }
    return variable.type === 'Secret' ? '•••••• Key Vault reference' : value;
}

const Value = ({ styles, variable, value, active = true }: { styles: Styles; variable: EnvironmentVariable; value: string | null; active?: boolean }) => {
    const text = displayValue(variable, value);
    if (text === null) {
        return <span className={styles.empty}>Not set</span>;
    }
    if (text === '') {
        return <span className={styles.empty}>Empty</span>;
    }
    return (
        <span className={active ? styles.primary : styles.secondary} title={variable.type === 'Secret' ? undefined : text}>
            {text}
        </span>
    );
};

interface EditableLineProps {
    styles: Styles;
    variable: EnvironmentVariable;
    kind: ValueKind;
    environmentName: string;
    onEdit: () => void;
    className?: string;
    children: React.ReactNode;
}

/** A value with a pencil (shown on cell hover) that opens the details dialog straight in edit mode. */
const EditableLine = ({ styles, variable, kind, environmentName, onEdit, className, children }: EditableLineProps) => (
    <div className={mergeClasses(styles.editLine, className)}>
        <span className={styles.editLineText}>{children}</span>
        {variable.type !== 'Secret' && (
            <button
                type="button"
                data-edit-button
                // Same tooltip text as the pencils in the details dialog; the label adds the environment for screen readers
                data-tooltip={`Edit ${kind} value`}
                className={styles.editButton}
                aria-label={`Edit ${kind} value in ${environmentName}`}
                onClick={(e) => {
                    // Don't let the row's own click (open details) fire as well
                    e.stopPropagation();
                    onEdit();
                }}
                onKeyDown={(e) => e.stopPropagation()}
            />
        )}
    </div>
);

/** Compare-mode cell: current value on top, default underneath; the one Dataverse doesn't use is greyed out. */
const EnvironmentCell = ({
    styles,
    variable,
    environmentName,
    onEdit,
    onCopy,
}: {
    styles: Styles;
    variable: EnvironmentVariable | null;
    environmentName: string;
    onEdit: (kind: ValueKind) => void;
    /** Set when the missing variable can be copied here from another environment. */
    onCopy?: () => void;
}) => {
    if (!variable) {
        return (
            <div className={styles.editLine}>
                <span className={mergeClasses(styles.editLineText, styles.empty)}>Not in this environment</span>
                {onCopy && (
                    <button
                        type="button"
                        data-copy-button
                        className={styles.copyButton}
                        aria-label={`Copy to ${environmentName}`}
                        onClick={(e) => {
                            // Don't let the row's own click (open details) fire as well
                            e.stopPropagation();
                            onCopy();
                        }}
                        onKeyDown={(e) => e.stopPropagation()}
                    >
                        Copy here
                    </button>
                )}
            </div>
        );
    }
    const usesCurrent = variable.currentValue !== null;
    return (
        <div className={styles.lines}>
            <EditableLine styles={styles} variable={variable} kind="current" environmentName={environmentName} onEdit={() => onEdit('current')}>
                <Value styles={styles} variable={variable} value={variable.currentValue} />
            </EditableLine>
            <EditableLine
                styles={styles}
                variable={variable}
                kind="default"
                environmentName={environmentName}
                onEdit={() => onEdit('default')}
                className={mergeClasses(styles.small, styles.secondary)}
            >
                Default: <Value styles={styles} variable={variable} value={variable.defaultValue} active={!usesCurrent} />
            </EditableLine>
        </div>
    );
};

interface VariableTableRowProps {
    row: VariableRow;
    snapshots: EnvironmentSnapshot[];
    onSelect: (row: VariableRow, edit?: EditTarget) => void;
    onCopy: (row: VariableRow, slot: number) => void;
}

/**
 * Memoised so rows that stay in the result set don't re-render while the user types a search.
 * Plain <tr>/<td> (styled like the header's Fluent cells) because Fluent's TableRow/TableCell add
 * per-cell work that made filtering hundreds of rows noticeably slow.
 */
const VariableTableRow = React.memo(({ row, snapshots, onSelect, onCopy }: VariableTableRowProps) => {
    const styles = useStyles();
    const isComparing = snapshots.length > 1;
    const single = row.values[0];
    const cell = mergeClasses(styles.cell, isComparing && styles.tallCell);

    const envCellClass = (variable: EnvironmentVariable | null) => {
        if (!variable) return mergeClasses(cell, styles.missingCell);
        return mergeClasses(cell, row.status === 'different' && styles.differentCell);
    };

    return (
        <tr className={styles.row} onClick={() => onSelect(row)} onKeyDown={(e) => e.key === 'Enter' && onSelect(row)} tabIndex={0}>
            <td className={mergeClasses(cell, styles.name)} title={row.displayName}>
                {row.displayName}
            </td>
            <td className={mergeClasses(cell, styles.secondary)} title={row.schemaName}>
                {row.schemaName}
            </td>
            <td className={cell}>
                <TypePill type={row.type} />
            </td>
            {isComparing ? (
                <>
                    <td className={cell}>
                        <StatusPill status={row.status} />
                    </td>
                    {row.values.map((variable, i) => (
                        <td key={snapshots[i].slot} className={mergeClasses(envCellClass(variable), styles.valueCell)}>
                            {snapshots[i].error ? (
                                <span className={styles.empty}>Failed to load</span>
                            ) : (
                                <EnvironmentCell
                                    styles={styles}
                                    variable={variable}
                                    environmentName={snapshots[i].connection.name}
                                    onEdit={(kind) => onSelect(row, { slot: snapshots[i].slot, kind })}
                                    onCopy={canCopy(row.type) ? () => onCopy(row, snapshots[i].slot) : undefined}
                                />
                            )}
                        </td>
                    ))}
                </>
            ) : (
                single && (
                    <>
                        <td className={mergeClasses(cell, styles.valueCell)}>
                            <EditableLine
                                styles={styles}
                                variable={single}
                                kind="default"
                                environmentName={snapshots[0].connection.name}
                                onEdit={() => onSelect(row, { slot: snapshots[0].slot, kind: 'default' })}
                            >
                                <Value styles={styles} variable={single} value={single.defaultValue} active={single.currentValue === null} />
                            </EditableLine>
                        </td>
                        <td className={mergeClasses(cell, styles.valueCell)}>
                            <EditableLine
                                styles={styles}
                                variable={single}
                                kind="current"
                                environmentName={snapshots[0].connection.name}
                                onEdit={() => onSelect(row, { slot: snapshots[0].slot, kind: 'current' })}
                            >
                                <Value styles={styles} variable={single} value={single.currentValue} />
                            </EditableLine>
                        </td>
                    </>
                )
            )}
            <td className={styles.filler} aria-hidden />
        </tr>
    );
});

interface Column {
    id: string;
    label: React.ReactNode;
    icon?: React.ReactElement;
    title?: string;
    defaultWidth: number;
}

interface ColumnFilterButtonProps {
    columnId: string;
    label: string;
    filter: ColumnFilter | undefined;
    onChange: (id: string, filter: ColumnFilter | undefined) => void;
}

/** Header button opening "<column> [operator] [value]" — applies as you type. */
const ColumnFilterButton: React.FC<ColumnFilterButtonProps> = ({ columnId, label, filter, onChange }) => {
    const styles = useStyles();
    const [open, setOpen] = React.useState(false);
    const operator = filter?.operator ?? 'contains';
    const value = filter?.value ?? '';
    const active = isFilterActive(filter);

    const update = (next: Partial<ColumnFilter>) => onChange(columnId, { operator, value, ...next });

    return (
        <Popover open={open} onOpenChange={(_e, data) => setOpen(data.open)} positioning="below-end" trapFocus>
            <PopoverTrigger disableButtonEnhancement>
                <Tooltip content={active ? `Filtered: ${label} ${FILTER_OPERATORS.find((o) => o.value === operator)?.label} ${operatorNeedsValue(operator) ? `"${value}"` : ''}` : `Filter ${label}`} relationship="label">
                    <Button
                        size="small"
                        appearance="subtle"
                        className={mergeClasses(styles.filterButton, active && styles.filterButtonActive)}
                        icon={active ? <Filter16Filled /> : <Filter16Regular />}
                    />
                </Tooltip>
            </PopoverTrigger>
            <PopoverSurface className={styles.filterSurface} aria-label={`Filter ${label}`}>
                <Text weight="semibold" size={200}>
                    {label}
                </Text>
                <Select size="small" aria-label="Condition" value={operator} onChange={(_e, data) => update({ operator: data.value as FilterOperator })}>
                    {FILTER_OPERATORS.map((o) => (
                        <option key={o.value} value={o.value}>
                            {o.label}
                        </option>
                    ))}
                </Select>
                {operatorNeedsValue(operator) && (
                    <Input
                        {...NO_TEXT_ASSIST}
                        size="small"
                        aria-label="Value"
                        placeholder="Value"
                        autoFocus
                        value={value}
                        onChange={(_e, data) => update({ value: data.value })}
                        onKeyDown={(e) => e.key === 'Enter' && setOpen(false)}
                    />
                )}
                <div className={styles.filterActions}>
                    <Button
                        size="small"
                        appearance="subtle"
                        disabled={!filter}
                        onClick={() => {
                            onChange(columnId, undefined);
                            setOpen(false);
                        }}
                    >
                        Clear
                    </Button>
                    <Button size="small" appearance="primary" onClick={() => setOpen(false)}>
                        Done
                    </Button>
                </div>
            </PopoverSurface>
        </Popover>
    );
};

interface ResizableHeaderCellProps {
    column: Column;
    width: number;
    onResize: (id: string, width: number | undefined) => void;
    filter: ColumnFilter | undefined;
    onFilterChange: (id: string, filter: ColumnFilter | undefined) => void;
}

/** Header cell with a drag handle on its right edge. Double-click the handle to reset; arrow keys resize too. */
const ResizableHeaderCell: React.FC<ResizableHeaderCellProps> = ({ column, width, onResize, filter, onFilterChange }) => {
    const styles = useStyles();
    const [isResizing, setIsResizing] = React.useState(false);
    const drag = useRef<{ startX: number; startWidth: number } | null>(null);

    const onPointerDown = useCallback(
        (e: React.PointerEvent<HTMLDivElement>) => {
            e.preventDefault();
            e.stopPropagation();
            e.currentTarget.setPointerCapture(e.pointerId);
            drag.current = { startX: e.clientX, startWidth: width };
            setIsResizing(true);
        },
        [width]
    );

    const onPointerMove = useCallback(
        (e: React.PointerEvent<HTMLDivElement>) => {
            if (!drag.current) return;
            onResize(column.id, drag.current.startWidth + e.clientX - drag.current.startX);
        },
        [column.id, onResize]
    );

    const endDrag = useCallback(() => {
        drag.current = null;
        setIsResizing(false);
    }, []);

    const onKeyDown = useCallback(
        (e: React.KeyboardEvent<HTMLDivElement>) => {
            const step = e.shiftKey ? 64 : 16;
            if (e.key === 'ArrowLeft') onResize(column.id, width - step);
            else if (e.key === 'ArrowRight') onResize(column.id, width + step);
            else if (e.key === 'Home') onResize(column.id, undefined);
            else return;
            e.preventDefault();
        },
        [column.id, width, onResize]
    );

    return (
        <TableHeaderCell className={mergeClasses(styles.headerCell, styles.resizableHeader)}>
            <span className={styles.headerInner}>
                <span className={mergeClasses(styles.headerContent, styles.headerLabel)} title={column.title}>
                    {column.icon}
                    <span className={styles.truncate}>{column.label}</span>
                </span>
                <ColumnFilterButton
                    columnId={column.id}
                    label={typeof column.label === 'string' ? column.label : column.id}
                    filter={filter}
                    onChange={onFilterChange}
                />
            </span>
            <div
                role="separator"
                aria-orientation="vertical"
                aria-label="Resize column"
                aria-valuenow={width}
                aria-valuemin={MIN_COLUMN_WIDTH}
                tabIndex={0}
                title="Drag to resize, double-click to reset"
                className={mergeClasses(styles.resizeHandle, isResizing && styles.resizing)}
                onPointerDown={onPointerDown}
                onPointerMove={onPointerMove}
                onPointerUp={endDrag}
                onPointerCancel={endDrag}
                onDoubleClick={() => onResize(column.id, undefined)}
                onKeyDown={onKeyDown}
                onClick={(e) => e.stopPropagation()}
            />
        </TableHeaderCell>
    );
};

interface VariablesTableProps {
    rows: VariableRow[];
    snapshots: EnvironmentSnapshot[];
    /** Opens the details dialog; with an edit target it opens straight into that value's editor. */
    onSelect: (row: VariableRow, edit?: EditTarget) => void;
    /** Opens the copy dialog for a variable missing from an environment. */
    onCopy: (row: VariableRow, slot: number) => void;
    columnFilters: Record<string, ColumnFilter>;
    onColumnFilterChange: (id: string, filter: ColumnFilter | undefined) => void;
    /** Shown under the header when no rows match, so the column filters stay editable. */
    emptyState?: React.ReactNode;
}

export const VariablesTable: React.FC<VariablesTableProps> = ({ rows, snapshots, onSelect, onCopy, columnFilters, onColumnFilterChange, emptyState }) => {
    const styles = useStyles();
    const { widths, setWidth } = useColumnWidths();
    const isComparing = snapshots.length > 1;

    const columns: Column[] = [
        { id: 'displayName', label: 'Display name', icon: <TextT16Regular />, defaultWidth: 220 },
        { id: 'schemaName', label: 'Schema name', icon: <Code16Regular />, defaultWidth: 220 },
        { id: 'type', label: 'Type', icon: <Shapes16Regular />, defaultWidth: 120 },
        ...(isComparing
            ? [
                  { id: 'comparison', label: 'Comparison', icon: <BranchCompare16Regular />, defaultWidth: 120 },
                  ...snapshots.map((snapshot) => ({
                      id: `env${snapshot.slot}`,
                      label: snapshot.connection.name,
                      icon: <EnvDot color={snapshot.connection.environmentColor} />,
                      title: snapshot.connection.url,
                      defaultWidth: 280,
                  })),
              ]
            : [
                  { id: 'defaultValue', label: 'Default value', icon: <TextT16Regular />, defaultWidth: 360 },
                  { id: 'currentValue', label: 'Current value', icon: <TextT16Regular />, defaultWidth: 360 },
              ]),
    ];

    const widthOf = (column: Column) => widths[column.id] ?? column.defaultWidth;
    const totalWidth = columns.reduce((sum, column) => sum + widthOf(column), 0);

    return (
        <div className={styles.container}>
            {/* Fixed layout + <colgroup> drives widths; the trailing filler column soaks up spare space */}
            <Table className={styles.table} style={{ width: `max(100%, ${totalWidth}px)` }} aria-label="Environment variables">
                <colgroup>
                    {columns.map((column) => (
                        <col key={column.id} style={{ width: widthOf(column) }} />
                    ))}
                    <col />
                </colgroup>
                <TableHeader className={styles.header}>
                    <TableRow>
                        {columns.map((column) => (
                            <ResizableHeaderCell
                                key={column.id}
                                column={column}
                                width={widthOf(column)}
                                onResize={setWidth}
                                filter={columnFilters[column.id]}
                                onFilterChange={onColumnFilterChange}
                            />
                        ))}
                        <TableHeaderCell className={styles.fillerHeader} aria-hidden />
                    </TableRow>
                </TableHeader>
                <TableBody>
                    {rows.map((row) => (
                        <VariableTableRow key={row.key} row={row} snapshots={snapshots} onSelect={onSelect} onCopy={onCopy} />
                    ))}
                </TableBody>
            </Table>
            {rows.length === 0 && emptyState && <div className={styles.emptyState}>{emptyState}</div>}
        </div>
    );
};
