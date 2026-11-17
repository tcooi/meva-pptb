import React from 'react';
import { Text, makeStyles, mergeClasses, tokens } from '@fluentui/react-components';
import type { CompareStatus, VariableType } from '../services/environmentVariables';

/**
 * Turns off the browser's spell checking and autocorrect on text fields. PPTB leaves Electron's spell
 * checker on for tools, and its first use (loading the dictionaries) makes typing stutter; these fields hold
 * names, URLs and JSON anyway, where red squiggles and autocorrect only get in the way.
 */
export const NO_TEXT_ASSIST = { spellCheck: false, autoComplete: 'off', autoCorrect: 'off', autoCapitalize: 'off' } as const;

/** Must match features.connections.max in package.json. */
export const MAX_ENVIRONMENTS = 5;

const useStyles = makeStyles({
    pill: {
        display: 'inline-flex',
        alignItems: 'center',
        height: '20px',
        padding: '0 8px',
        borderRadius: tokens.borderRadiusCircular,
        fontSize: tokens.fontSizeBase200,
        lineHeight: '20px',
        whiteSpace: 'nowrap',
    },
    // Pastel "single select" colours, one per type
    String: { backgroundColor: tokens.colorPaletteBlueBackground2, color: tokens.colorPaletteBlueForeground2 },
    Number: { backgroundColor: tokens.colorPaletteTealBackground2, color: tokens.colorPaletteTealForeground2 },
    Boolean: { backgroundColor: tokens.colorPalettePurpleBackground2, color: tokens.colorPalettePurpleForeground2 },
    JSON: { backgroundColor: tokens.colorPaletteMarigoldBackground2, color: tokens.colorPaletteMarigoldForeground2 },
    'Data source': { backgroundColor: tokens.colorPaletteLightGreenBackground2, color: tokens.colorPaletteLightGreenForeground2 },
    Secret: { backgroundColor: tokens.colorPaletteBerryBackground2, color: tokens.colorPaletteBerryForeground2 },
    Unknown: { backgroundColor: tokens.colorNeutralBackground3, color: tokens.colorNeutralForeground3 },
    match: { backgroundColor: tokens.colorStatusSuccessBackground1, color: tokens.colorStatusSuccessForeground1 },
    different: { backgroundColor: tokens.colorStatusWarningBackground1, color: tokens.colorStatusWarningForeground3 },
    missing: { backgroundColor: tokens.colorStatusDangerBackground1, color: tokens.colorStatusDangerForeground1 },
    dot: {
        display: 'inline-block',
        width: '8px',
        height: '8px',
        borderRadius: '50%',
        flexShrink: 0,
    },
    empty: {
        flex: 1,
        display: 'flex',
        flexDirection: 'column',
        alignItems: 'center',
        justifyContent: 'center',
        gap: tokens.spacingVerticalXS,
        padding: tokens.spacingVerticalXXL,
        textAlign: 'center',
    },
    emptyIcon: {
        fontSize: '24px',
        color: tokens.colorNeutralForeground3,
        marginBottom: tokens.spacingVerticalXS,
    },
    slots: {
        display: 'flex',
        gap: '4px',
        margin: `${tokens.spacingVerticalXS} 0`,
    },
    // Rough sketch of PPTB's numbered connection squares, so the explanation is easy to map to the real footer
    slot: {
        display: 'inline-flex',
        alignItems: 'center',
        justifyContent: 'center',
        width: '20px',
        height: '20px',
        borderRadius: tokens.borderRadiusSmall,
        fontSize: tokens.fontSizeBase100,
        fontWeight: tokens.fontWeightSemibold,
        border: `1px solid ${tokens.colorNeutralStroke1}`,
        color: tokens.colorNeutralForeground3,
    },
    slotUsed: {
        backgroundColor: tokens.colorBrandBackground,
        border: `1px solid ${tokens.colorBrandBackground}`,
        color: tokens.colorNeutralForegroundOnBrand,
    },
    slotNext: {
        border: `1px dashed ${tokens.colorBrandStroke1}`,
        color: tokens.colorBrandForeground1,
    },
    hintText: {
        color: tokens.colorNeutralForeground2,
    },
    emptyDescription: {
        color: tokens.colorNeutralForeground3,
        maxWidth: '380px',
    },
});

export const TypePill: React.FC<{ type: VariableType }> = ({ type }) => {
    const styles = useStyles();
    return <span className={mergeClasses(styles.pill, styles[type])}>{type}</span>;
};

const STATUS_LABELS: Record<Exclude<CompareStatus, 'single'>, string> = {
    match: 'Same',
    different: 'Different',
    missing: 'Missing',
};

export const StatusPill: React.FC<{ status: CompareStatus }> = ({ status }) => {
    const styles = useStyles();
    if (status === 'single') return null;
    return <span className={mergeClasses(styles.pill, styles[status])}>{STATUS_LABELS[status]}</span>;
};

/** Small colour dot for an environment; falls back to the brand colour when PPTB has no colour set. */
export const EnvDot: React.FC<{ color?: string }> = ({ color }) => {
    const styles = useStyles();
    return <span className={styles.dot} style={{ backgroundColor: color || tokens.colorBrandBackground }} aria-hidden />;
};

export const EmptyState: React.FC<{ icon: React.ReactElement; title: string; description?: string }> = ({ icon, title, description }) => {
    const styles = useStyles();
    return (
        <div className={styles.empty}>
            <span className={styles.emptyIcon}>{icon}</span>
            <Text weight="semibold">{title}</Text>
            {description && (
                <Text size={200} className={styles.emptyDescription}>
                    {description}
                </Text>
            )}
        </div>
    );
};

/**
 * Explains how to add an environment. PPTB doesn't let tools open its connection picker,
 * so this points at the numbered squares in PPTB's footer instead.
 */
export const AddEnvironmentHelp: React.FC<{ used: number }> = ({ used }) => {
    const styles = useStyles();
    return (
        <>
            <Text weight="semibold">Compare another environment</Text>
            <Text size={200} className={styles.hintText}>
                Click the next empty numbered square in the Power Platform ToolBox footer at the bottom of the window, then choose a connection. The grid
                reloads with the new environment as an extra column.
            </Text>
            <span className={styles.slots} aria-hidden>
                {Array.from({ length: MAX_ENVIRONMENTS }, (_, i) => (
                    <span key={i} className={mergeClasses(styles.slot, i < used && styles.slotUsed, i === used && styles.slotNext)}>
                        {i + 1}
                    </span>
                ))}
            </span>
            <Text size={200} className={styles.hintText}>
                You can compare up to {MAX_ENVIRONMENTS} environments. To change or remove one, use Manage Connections in the footer.
            </Text>
        </>
    );
};
