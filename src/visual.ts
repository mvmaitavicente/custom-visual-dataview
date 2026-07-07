/// <reference path="./typings.d.ts" />

"use strict";

import powerbi from "powerbi-visuals-api";
import { FormattingSettingsService } from "powerbi-visuals-utils-formattingmodel";
import "./../style/visual.less";

import DataView = powerbi.DataView;
import DataViewMetadataColumn = powerbi.DataViewMetadataColumn;
import DataViewTable = powerbi.DataViewTable;
import DataViewTableRow = powerbi.DataViewTableRow;
import IVisual = powerbi.extensibility.visual.IVisual;
import IVisualHost = powerbi.extensibility.visual.IVisualHost;
import PrimitiveValue = powerbi.PrimitiveValue;
import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import VisualUpdateOptions = powerbi.extensibility.visual.VisualUpdateOptions;

import { VisualFormattingSettingsModel } from "./settings";

type RiskLevel = "Bajo" | "Medio" | "Alto" | "Muy Alto";

const riskLevels: RiskLevel[] = ["Bajo", "Medio", "Alto", "Muy Alto"];

interface ColumnIndexes {
    region: number;
    provincia: number;
    codigoLocal: number;
    nombreLocal: number;
    unidadGerencial: number;
    montoIntervencion: number;
    beneficiarios: number;
    montoAsignado: number;
    montoTransferencia: number;
    montoRetirado: number;
    grupo: number;
    cantidad: number;
    distrito: number;
    idSolicitud: number;
    estadoSolicitud: number;
    nivelRiesgo: number;
    montoInversion: number;
    latitud: number;
    longitud: number;
}

interface SourceRecord {
    unidadGerencial: string;
    region: string;
    provincia: string;
    codigoLocal: string;
    nombreLocal: string;
    montoIntervencion: number;
    beneficiarios: number;
    montoAsignado: number;
    montoTransferencia: number;
    montoRetirado: number;
    grupo: string;
    cantidad: number;
    distrito: string;
    idSolicitud: string;
    estadoSolicitud: string;
    nivelRiesgo: string;
    montoInversion: number;
    latitud: number;
    longitud: number;
}

interface RiskBuckets {
    bajo: number;
    medio: number;
    alto: number;
    muyAlto: number;
}

interface RiskSetBuckets {
    bajo: Set<string>;
    medio: Set<string>;
    alto: Set<string>;
    muyAlto: Set<string>;
}

interface MapRegionShape {
    name: string;
    path: string;
    labelX: number;
    labelY: number;
}

interface RegionAggregate {
    region: string;
    totalColegios: number;
    montoInversion: number;
    estudiantesBeneficiarios: number;
    numeroSolicitudes: number;
    risk: RiskLevel;
    predominantRisk: RiskLevel;
    weightedRisk: RiskLevel;
    weightedScore: number;
    riskBreakdown: RiskBuckets;
}

const peruRegionShapes: MapRegionShape[] = [
    { name: "TUMBES", path: "M120 45 L190 52 L180 95 L115 88 Z", labelX: 150, labelY: 72 },
    { name: "PIURA", path: "M100 92 L205 98 L218 170 L168 196 L112 170 Z", labelX: 158, labelY: 142 },
    { name: "LAMBAYEQUE", path: "M126 180 L205 178 L210 225 L135 230 Z", labelX: 170, labelY: 205 },
    { name: "CAJAMARCA", path: "M210 105 L305 115 L318 230 L214 225 L205 170 Z", labelX: 265, labelY: 170 },
    { name: "AMAZONAS", path: "M315 92 L430 115 L405 245 L318 230 L305 115 Z", labelX: 370, labelY: 170 },
    { name: "LORETO", path: "M430 80 L760 130 L790 320 L650 395 L520 345 L405 245 Z", labelX: 615, labelY: 235 },
    { name: "LA LIBERTAD", path: "M145 230 L250 232 L264 310 L170 318 Z", labelX: 205, labelY: 274 },
    { name: "SAN MARTIN", path: "M325 235 L420 252 L445 390 L338 375 L264 310 Z", labelX: 370, labelY: 315 },
    { name: "ANCASH", path: "M172 318 L280 314 L305 405 L220 440 L175 390 Z", labelX: 235, labelY: 375 },
    { name: "HUANUCO", path: "M305 405 L338 375 L452 395 L460 505 L345 510 Z", labelX: 390, labelY: 450 },
    { name: "UCAYALI", path: "M460 505 L452 395 L650 395 L705 595 L565 655 Z", labelX: 570, labelY: 520 },
    { name: "PASCO", path: "M345 510 L460 505 L475 575 L370 590 Z", labelX: 415, labelY: 548 },
    { name: "LIMA PROVINCIAS", path: "M220 440 L345 510 L370 590 L318 660 L225 585 Z", labelX: 285, labelY: 555 },
    { name: "LIMA METROPOLITANA", path: "M212 590 L246 600 L238 635 L205 625 Z", labelX: 226, labelY: 615 },
    { name: "CALLAO", path: "M190 600 L210 604 L205 628 L185 622 Z", labelX: 198, labelY: 616 },
    { name: "JUNIN", path: "M370 590 L475 575 L510 675 L400 710 L318 660 Z", labelX: 425, labelY: 642 },
    { name: "MADRE DE DIOS", path: "M650 600 L790 560 L842 705 L700 740 L565 655 Z", labelX: 710, labelY: 660 },
    { name: "CUSCO", path: "M510 675 L565 655 L700 740 L642 860 L500 802 Z", labelX: 595, labelY: 750 },
    { name: "HUANCAVELICA", path: "M290 662 L400 710 L392 790 L290 760 Z", labelX: 340, labelY: 735 },
    { name: "ICA", path: "M225 585 L290 662 L290 760 L240 790 L192 700 Z", labelX: 250, labelY: 695 },
    { name: "AYACUCHO", path: "M392 790 L400 710 L500 802 L470 900 L360 858 Z", labelX: 430, labelY: 810 },
    { name: "APURIMAC", path: "M470 900 L500 802 L595 868 L550 940 Z", labelX: 520, labelY: 880 },
    { name: "AREQUIPA", path: "M240 790 L360 858 L470 900 L428 1040 L275 980 Z", labelX: 350, labelY: 930 },
    { name: "PUNO", path: "M595 868 L642 860 L742 980 L660 1090 L550 940 Z", labelX: 635, labelY: 950 },
    { name: "MOQUEGUA", path: "M428 1040 L550 940 L575 1035 L500 1100 Z", labelX: 500, labelY: 1018 },
    { name: "TACNA", path: "M500 1100 L575 1035 L660 1090 L610 1160 L510 1165 Z", labelX: 570, labelY: 1120 }
];

interface MetricBucket {
    key: string;
    label: string;
    filas: number;
    colegios: Set<string>;
    solicitudes: Set<string>;
    regiones: Set<string>;
    provincias: Set<string>;
    distritos: Set<string>;
    unidades: Set<string>;
    montoTotal: number;
    beneficiariosTotal: number;
    riesgo: {
        bajo: number;
        medio: number;
        alto: number;
        muyAlto: number;
    };
    riesgoColegios: RiskSetBuckets;
}

interface AnalyticsResult {
    totals: MetricBucket;
    byRegion: Map<string, MetricBucket>;
    byProvincia: Map<string, MetricBucket>;
    byDistrito: Map<string, MetricBucket>;
    byUnidad: Map<string, MetricBucket>;
    byEstado: Map<string, MetricBucket>;
    byRiesgo: Map<string, MetricBucket>;
    byColegio: Map<string, MetricBucket>;
    byUnidadKpi: Map<string, UnitKpiBucket>;
    provinciasByRegion: Map<string, Set<string>>;
    distritosByProvincia: Map<string, Set<string>>;
    colegiosByDistrito: Map<string, Set<string>>;
}

interface BucketSummary {
    key: string;
    label: string;
    level?: SummaryLevel;
    filas: number;
    colegiosUnicos: number;
    solicitudesUnicas: number;
    regionesUnicas: number;
    provinciasUnicas: number;
    distritosUnicos: number;
    unidadesUnicas: number;
    montoTotal: number;
    beneficiariosTotal: number;
    riesgo: RiskBuckets;
    riesgoColegios: RiskBuckets;
    bajo: number;
    medio: number;
    alto: number;
    muyAlto: number;
    porcentajeCritico: number;
    scoreRiesgo: number;
    clasificacion: RiskLevel;
}

interface RiskLevelSummary extends BucketSummary {
    porcentajeSobreTotal: number;
}

type UnitKpiName = "UGRD" | "UGEO" | "UGSC" | "UGM" | "UGME";
type VisualView = "summary" | "ugme";

interface GroupMetricBucket {
    grupoKey: string;
    grupoLabel: string;
    montoIntervencion: number;
    beneficiarios: number;
    cantidad: number;
    colegios: Set<string>;
    regiones: Set<string>;
    filas: number;
}

interface GroupMetricSummary {
    grupo: string;
    montoIntervencion: number;
    beneficiarios: number;
    cantidad: number;
    colegios: number;
    regiones: number;
    filas: number;
}

interface UnitKpiBucket {
    unitKey: string;
    label: string;
    filas: number;
    regiones: Set<string>;
    colegios: Set<string>;
    solicitudes: Set<string>;
    beneficiarios: number;
    montoIntervencion: number;
    montoAsignado: number;
    montoTransferencia: number;
    montoRetirado: number;
    solicitudesRevision: Set<string>;
    solicitudesCulminadas: Set<string>;
    colegiosConTransferencia: Set<string>;
    colegiosConRetiro: Set<string>;
    grupos: Map<string, GroupMetricBucket>;
}

interface UnitKpiSummary {
    unit: UnitKpiName;
    unitKey: UnitKpiName;
    unidad: string;
    label: string;
    filas: number;
    solicitudes: number;
    regiones: number;
    colegios: number;
    montoIntervencion: number;
    montoAsignado: number;
    montoTransferencia: number;
    montoRetirado: number;
    beneficiarios: number;
    solicitudesRevision: number;
    solicitudesCulminadas: number;
    colegiosConTransferencia: number;
    colegiosConRetiro: number;
    grupos: GroupMetricSummary[];
    registros: number;
}

interface UgmeKpiSummary {
    unitKey: "UGME";
    regiones: number;
    colegios: number;
    montoIntervencion: number;
    grupos: GroupMetricSummary[];
}

interface UnitTheme {
    color: string;
    background: string;
    icon: string;
    name: string;
}

interface UnitKpiCell {
    icon: string;
    label: string;
    value: string;
}

interface UnitDetailRow {
    unit: UnitKpiName;
    codigoLocal: string;
    nombreLocal: string;
    region: string;
    provincia: string;
    distrito: string;
    grupo?: string;
    montoIntervencion: number;
    beneficiarios: number;
    cantidad: number;
    solicitudes: Set<string>;
    solicitudesRevision: Set<string>;
    solicitudesCulminadas: Set<string>;
    montoAsignado: number;
    montoTransferencia: number;
    montoRetirado: number;
}

interface DetailColumn {
    label: string;
    value: (row: UnitDetailRow) => string;
}

interface RegionTooltipSummary {
    colegios: number;
    provincias: number;
}

const ugmeGrupoOrder = [
    "MOBILIARIO",
    "EQUIPAMIENTO",
    "LABORATORIOS",
    "TALLERES",
    "ESCUELA MODULAR",
    "AULA MODULAR",
    "KIT DE PARARRAYO",
    "MODULO SS.HH.",
    "REDES COMPLEMENTARIAS"
];

type SummaryLevel = "nacional" | "region" | "provincia" | "distrito" | "colegio";

interface LazyNavigationDiagnostics {
    regionesIndexadas: number;
    provinciasIndexadas: number;
    distritosIndexados: number;
    colegiosIndexados: number;
    relacionesProvinciasByRegion: number;
    relacionesDistritosByProvincia: number;
    relacionesColegiosByDistrito: number;
    primeraRegion: string;
    provinciasPrimeraRegion: number;
    primerasProvincias: string[][];
    primeraProvincia: string;
    distritosPrimeraProvincia: number;
    primerosDistritos: string[][];
    primerDistrito: string;
    colegiosPrimerDistrito: number;
    primerosColegios: string[][];
    notaProvinciaDistrito: string;
}

interface RecordKeys {
    regionKey: string;
    provinciaKey: string;
    distritoKey: string;
    unidadKey: string;
    estadoKey: string;
    riesgoKey: string;
    colegioKey: string;
    solicitudKey: string;
    regionLabel: string;
    provinciaLabel: string;
    distritoLabel: string;
    unidadLabel: string;
    estadoLabel: string;
    riesgoLabel: RiskLevel;
    colegioLabel: string;
}

interface PerformanceMetrics {
    updateTotalMs: number;
    segmentLoadMs: number;
    readDataViewMs: number;
    findColumnsMs: number;
    buildRecordsMs: number;
    accumulateRowsMs: number;
    buildAnalyticsEngineMs: number;
    buildFlatIndexesMs: number;
    buildRelationIndexesMs: number;
    lazySampleQueryMs: number;
    renderMs: number;
    indexedRegiones: number;
    indexedProvincias: number;
    indexedDistritos: number;
    indexedColegios: number;
    fetchCount: number;
    rowsPerSecond: number;
    msPer10kRows: number;
    analyticsCacheHit: boolean;
    analyticsRebuilt: boolean;
    datasetSignature: string;
    fetchWaitMs?: number;
}

interface TableDiagnostics {
    dataViewCount: number;
    hasTable: boolean;
    columnIndexes: ColumnIndexes;
    columns: DataViewMetadataColumn[];
    rows: DataViewTableRow[];
    records: SourceRecord[];
    segmentFilas: number;
    accumulatedFilas: number;
    hasMoreData: boolean;
    isLoading: boolean;
    fetchCount: number;
    datasetSignature: string;
    analyticsCacheHit: boolean;
    analyticsRebuilt: boolean;
    dataReductionText: string;
    analytics: AnalyticsEngine | null;
    performanceMetrics: PerformanceMetrics;
}

function measure<T>(label: string, fn: () => T): { value: T; ms: number } {
    const start = performance.now();
    const value = fn();
    const end = performance.now();
    console.log(`${label} ms`, end - start);
    return { value, ms: end - start };
}

function getRiskRatioColor(ratio: number): string {
    if (!Number.isFinite(ratio) || ratio < 0) {
        return "#d1d5db";
    }

    const clamped = Math.max(0, Math.min(1, ratio));

    if (clamped <= 0.2) {
        return "#86efac";
    }

    if (clamped <= 0.4) {
        return "#d9f99d";
    }

    if (clamped <= 0.6) {
        return "#fde68a";
    }

    if (clamped <= 0.8) {
        return "#fdba74";
    }

    return "#ef4444";
}

class AnalyticsEngine {
    private result: AnalyticsResult;
    private flatIndexesMs = 0;
    private relationIndexesMs = 0;
    private lazySampleQueryMs = 0;
    private regionSummaries: BucketSummary[] = [];
    private unitSummaries: BucketSummary[] = [];
    private stateSummaries: BucketSummary[] = [];
    private riskSummaries: RiskLevelSummary[] = [];

    private constructor(result: AnalyticsResult) {
        this.result = result;
    }

    public static build(records: SourceRecord[]): AnalyticsEngine {
        const result = AnalyticsEngine.createResult();
        let flatIndexesMs = 0;
        let relationIndexesMs = 0;

        for (const record of records) {
            const recordKeys = AnalyticsEngine.getRecordKeys(record);

            const flatStart = performance.now();
            AnalyticsEngine.updateBucket(result.totals, recordKeys, record);
            AnalyticsEngine.updateMapBucket(result.byRegion, recordKeys.regionKey, recordKeys.regionLabel, recordKeys, record);
            AnalyticsEngine.updateMapBucket(result.byProvincia, recordKeys.provinciaKey, recordKeys.provinciaLabel, recordKeys, record);
            AnalyticsEngine.updateMapBucket(result.byDistrito, recordKeys.distritoKey, recordKeys.distritoLabel, recordKeys, record);
            AnalyticsEngine.updateMapBucket(result.byUnidad, recordKeys.unidadKey, recordKeys.unidadLabel, recordKeys, record);
            AnalyticsEngine.updateMapBucket(result.byEstado, recordKeys.estadoKey, recordKeys.estadoLabel, recordKeys, record);
            AnalyticsEngine.updateMapBucket(result.byRiesgo, recordKeys.riesgoKey, recordKeys.riesgoLabel, recordKeys, record);
            AnalyticsEngine.updateMapBucket(result.byColegio, recordKeys.colegioKey, recordKeys.colegioLabel, recordKeys, record);
            AnalyticsEngine.updateUnitKpiBucket(result.byUnidadKpi, record);
            flatIndexesMs += performance.now() - flatStart;

            const relationStart = performance.now();
            AnalyticsEngine.addRelation(result.provinciasByRegion, recordKeys.regionKey, recordKeys.provinciaKey);
            AnalyticsEngine.addRelation(result.distritosByProvincia, recordKeys.provinciaKey, recordKeys.distritoKey);
            AnalyticsEngine.addRelation(result.colegiosByDistrito, recordKeys.distritoKey, recordKeys.colegioKey);
            relationIndexesMs += performance.now() - relationStart;
        }

        const engine = new AnalyticsEngine(result);
        engine.flatIndexesMs = flatIndexesMs;
        engine.relationIndexesMs = relationIndexesMs;
        engine.buildFlatSummaries();
        engine.measureLazySampleQuery();

        return engine;
    }

    public buildFlatSummaries(): void {
        this.regionSummaries = this.mapToSummaries(this.result.byRegion, "region");
        this.unitSummaries = this.mapToSummaries(this.result.byUnidad);
        this.stateSummaries = this.mapToSummaries(this.result.byEstado);
        this.riskSummaries = this.buildRiskSummaries();
    }

    public getResult(): AnalyticsResult {
        return this.result;
    }

    public getBuildMetrics(): Pick<PerformanceMetrics, "buildFlatIndexesMs" | "buildRelationIndexesMs" | "lazySampleQueryMs"> {
        return {
            buildFlatIndexesMs: this.flatIndexesMs,
            buildRelationIndexesMs: this.relationIndexesMs,
            lazySampleQueryMs: this.lazySampleQueryMs
        };
    }

    public getTotals(): BucketSummary {
        return this.bucketToSummary(this.result.totals, "nacional");
    }

    public getRegions(): BucketSummary[] {
        return this.regionSummaries;
    }

    public getProvincias(regionKey: string): BucketSummary[] {
        return this.relatedSummaries(this.result.provinciasByRegion, this.result.byProvincia, regionKey, "provincia");
    }

    public getDistritos(provinciaKey: string): BucketSummary[] {
        return this.relatedSummaries(this.result.distritosByProvincia, this.result.byDistrito, provinciaKey, "distrito");
    }

    public getUnidades(): BucketSummary[] {
        return this.unitSummaries;
    }

    public getEstados(): BucketSummary[] {
        return this.stateSummaries;
    }

    public getRiesgos(): RiskLevelSummary[] {
        return this.riskSummaries;
    }

    public getUnitKpis(): UnitKpiSummary[] {
        return [
            this.unitKpiBucketToSummary(this.result.byUnidadKpi.get("UGRD"), "UGRD"),
            this.unitKpiBucketToSummary(this.result.byUnidadKpi.get("UGEO"), "UGEO"),
            this.unitKpiBucketToSummary(this.result.byUnidadKpi.get("UGSC"), "UGSC"),
            this.unitKpiBucketToSummary(this.result.byUnidadKpi.get("UGM"), "UGM"),
            this.unitKpiBucketToSummary(this.result.byUnidadKpi.get("UGME"), "UGME")
        ];
    }

    public getUnitKpi(unitKey: UnitKpiName): UnitKpiSummary {
        const normalizedUnit = AnalyticsEngine.normalizeLookupKey(unitKey).toUpperCase() as UnitKpiName;
        return this.unitKpiBucketToSummary(this.result.byUnidadKpi.get(normalizedUnit), normalizedUnit);
    }

    public getUnitSummary(unitKey: string): UnitKpiSummary {
        return this.getUnitKpi(AnalyticsEngine.normalizeLookupKey(unitKey).toUpperCase() as UnitKpiName);
    }

    public getUgmeKpi(): UgmeKpiSummary {
        const bucket = this.result.byUnidadKpi.get("UGME");
        const summary = this.unitKpiBucketToSummary(bucket, "UGME");

        return {
            unitKey: "UGME",
            regiones: summary.regiones,
            colegios: summary.colegios,
            montoIntervencion: summary.montoIntervencion,
            grupos: summary.grupos
        };
    }

    public getRegionTooltip(regionKey: string): RegionTooltipSummary {
        const bucket = this.result.byRegion.get(regionKey) || this.result.byRegion.get(AnalyticsEngine.normalizeLookupKey(regionKey));

        return {
            colegios: bucket?.colegios.size || 0,
            provincias: bucket?.provincias.size || 0
        };
    }

    public getColegiosByDistrito(distritoKey: string): BucketSummary[] {
        return this.relatedSummaries(this.result.colegiosByDistrito, this.result.byColegio, distritoKey, "colegio");
    }

    public getBucketByRegion(regionKey: string): BucketSummary | undefined {
        return this.getBucketSummary(this.result.byRegion, regionKey, "region");
    }

    public getBucketByProvincia(provinciaKey: string): BucketSummary | undefined {
        return this.getBucketSummary(this.result.byProvincia, provinciaKey, "provincia");
    }

    public getBucketByDistrito(distritoKey: string): BucketSummary | undefined {
        return this.getBucketSummary(this.result.byDistrito, distritoKey, "distrito");
    }

    public getBucketByColegio(codigoLocal: string): BucketSummary | undefined {
        return this.getBucketSummary(this.result.byColegio, codigoLocal, "colegio");
    }

    public getLazyNavigationDiagnostics(): LazyNavigationDiagnostics {
        const firstRegion = Array.from(this.result.byRegion.keys())[0] || "";
        const firstProvincia = this.firstRelatedKey(this.result.provinciasByRegion, firstRegion);
        const firstDistrito = firstProvincia ? this.firstRelatedKey(this.result.distritosByProvincia, firstProvincia) : "";

        return {
            regionesIndexadas: this.result.byRegion.size,
            provinciasIndexadas: this.result.byProvincia.size,
            distritosIndexados: this.result.byDistrito.size,
            colegiosIndexados: this.result.byColegio.size,
            relacionesProvinciasByRegion: this.result.provinciasByRegion.size,
            relacionesDistritosByProvincia: this.result.distritosByProvincia.size,
            relacionesColegiosByDistrito: this.result.colegiosByDistrito.size,
            primeraRegion: this.result.byRegion.get(firstRegion)?.label || "(sin region)",
            provinciasPrimeraRegion: this.result.provinciasByRegion.get(firstRegion)?.size || 0,
            primerasProvincias: firstRegion
                ? this.relatedRows(this.result.provinciasByRegion, this.result.byProvincia, firstRegion, 5)
                : [],
            primeraProvincia: firstProvincia ? this.result.byProvincia.get(firstProvincia)?.label || firstProvincia : "(sin provincia)",
            distritosPrimeraProvincia: firstProvincia ? this.result.distritosByProvincia.get(firstProvincia)?.size || 0 : 0,
            primerosDistritos: firstProvincia
                ? this.relatedRows(this.result.distritosByProvincia, this.result.byDistrito, firstProvincia, 5)
                : [],
            primerDistrito: firstDistrito ? this.result.byDistrito.get(firstDistrito)?.label || firstDistrito : "(sin distrito)",
            colegiosPrimerDistrito: firstDistrito ? this.result.colegiosByDistrito.get(firstDistrito)?.size || 0 : 0,
            primerosColegios: firstDistrito
                ? this.relatedRows(this.result.colegiosByDistrito, this.result.byColegio, firstDistrito, 5)
                : [],
            notaProvinciaDistrito: "TODO: En produccion usar IdProvincia/IdDistrito reales para evitar cardinalidades artificiales."
        };
    }

    private static createResult(): AnalyticsResult {
        return {
            totals: AnalyticsEngine.createBucket("__totals__", "TOTAL"),
            byRegion: new Map<string, MetricBucket>(),
            byProvincia: new Map<string, MetricBucket>(),
            byDistrito: new Map<string, MetricBucket>(),
            byUnidad: new Map<string, MetricBucket>(),
            byEstado: new Map<string, MetricBucket>(),
            byRiesgo: new Map<string, MetricBucket>(),
            byColegio: new Map<string, MetricBucket>(),
            byUnidadKpi: new Map<string, UnitKpiBucket>(),
            provinciasByRegion: new Map<string, Set<string>>(),
            distritosByProvincia: new Map<string, Set<string>>(),
            colegiosByDistrito: new Map<string, Set<string>>()
        };
    }

    private static getRecordKeys(record: SourceRecord): RecordKeys {
        const riesgoLabel = AnalyticsEngine.normalizeRiskLevel(record.nivelRiesgo);
        const regionKey = AnalyticsEngine.normalizeKey(record.region);
        const provinciaBaseKey = AnalyticsEngine.normalizeKey(record.provincia);
        const provinciaKey = `${regionKey}|${provinciaBaseKey}`;
        const distritoBaseKey = AnalyticsEngine.normalizeKey(record.distrito);
        const distritoKey = `${provinciaKey}|${distritoBaseKey}`;

        // TODO: En produccion usar IdProvincia/IdDistrito reales para evitar cardinalidades artificiales.
        return {
            regionKey,
            provinciaKey,
            distritoKey,
            unidadKey: AnalyticsEngine.normalizeKey(record.unidadGerencial),
            estadoKey: AnalyticsEngine.normalizeKey(record.estadoSolicitud),
            riesgoKey: riesgoLabel,
            colegioKey: AnalyticsEngine.normalizeKey(record.codigoLocal),
            solicitudKey: AnalyticsEngine.normalizeKey(record.idSolicitud),
            regionLabel: AnalyticsEngine.displayValue(record.region),
            provinciaLabel: AnalyticsEngine.displayValue(record.provincia),
            distritoLabel: AnalyticsEngine.displayValue(record.distrito),
            unidadLabel: AnalyticsEngine.displayValue(record.unidadGerencial),
            estadoLabel: AnalyticsEngine.displayValue(record.estadoSolicitud),
            riesgoLabel,
            colegioLabel: AnalyticsEngine.displayValue(record.codigoLocal)
        };
    }

    private static normalizeKey(value: PrimitiveValue): string {
        return AnalyticsEngine.displayValue(value)
            .normalize("NFD")
            .replace(/[\u0300-\u036f]/g, "")
            .trim()
            .toLowerCase();
    }

    private static normalizeText(value: PrimitiveValue): string {
        return AnalyticsEngine.displayValue(value)
            .normalize("NFD")
            .replace(/[\u0300-\u036f]/g, "")
            .trim()
            .replace(/\s+/g, " ")
            .toUpperCase();
    }

    private static isRevisionState(value: string): boolean {
        return value === "EN REVISION" || value === "REVISION";
    }

    private static isCulminadoState(value: string): boolean {
        return value === "CULMINADO" || value === "CULMINADA" || value === "CULMINADOS" || value === "CULMINADAS";
    }

    private static normalizeGrupo(value: PrimitiveValue): string {
        const normalized = AnalyticsEngine.normalizeText(value)
            .replace(/[.]/g, "")
            .replace(/\s+/g, " ");

        if (normalized === "MODULO SSHH" || normalized === "MODULO DE SSHH" || normalized === "MODULO SS HH" || normalized === "MODULO DE SS HH") {
            return "MODULO SS.HH.";
        }

        return ugmeGrupoOrder.includes(normalized) ? normalized : "";
    }

    private static createBucket(key: string, label: string): MetricBucket {
        return {
            key,
            label,
            filas: 0,
            colegios: new Set<string>(),
            solicitudes: new Set<string>(),
            regiones: new Set<string>(),
            provincias: new Set<string>(),
            distritos: new Set<string>(),
            unidades: new Set<string>(),
            montoTotal: 0,
            beneficiariosTotal: 0,
            riesgo: {
                bajo: 0,
                medio: 0,
                alto: 0,
                muyAlto: 0
            },
            riesgoColegios: {
                bajo: new Set<string>(),
                medio: new Set<string>(),
                alto: new Set<string>(),
                muyAlto: new Set<string>()
            }
        };
    }

    private static updateMapBucket(
        map: Map<string, MetricBucket>,
        key: string,
        label: string,
        recordKeys: RecordKeys,
        record: SourceRecord
    ): void {
        const bucket = map.get(key) || AnalyticsEngine.createBucket(key, label);
        AnalyticsEngine.updateBucket(bucket, recordKeys, record);
        map.set(key, bucket);
    }

    private static updateBucket(bucket: MetricBucket, recordKeys: RecordKeys, record: SourceRecord): void {
        bucket.filas += 1;
        bucket.colegios.add(recordKeys.colegioKey);
        bucket.solicitudes.add(recordKeys.solicitudKey);
        bucket.regiones.add(recordKeys.regionKey);
        bucket.provincias.add(recordKeys.provinciaKey);
        bucket.distritos.add(recordKeys.distritoKey);
        bucket.unidades.add(recordKeys.unidadKey);
        bucket.montoTotal += record.montoInversion;
        bucket.beneficiariosTotal += Number(record.beneficiarios) || 0;

        if (recordKeys.riesgoLabel === "Bajo") {
            bucket.riesgo.bajo += 1;
            bucket.riesgoColegios.bajo.add(recordKeys.colegioKey);
        } else if (recordKeys.riesgoLabel === "Medio") {
            bucket.riesgo.medio += 1;
            bucket.riesgoColegios.medio.add(recordKeys.colegioKey);
        } else if (recordKeys.riesgoLabel === "Alto") {
            bucket.riesgo.alto += 1;
            bucket.riesgoColegios.alto.add(recordKeys.colegioKey);
        } else {
            bucket.riesgo.muyAlto += 1;
            bucket.riesgoColegios.muyAlto.add(recordKeys.colegioKey);
        }
    }

    private static updateUnitKpiBucket(
        map: Map<string, UnitKpiBucket>,
        record: SourceRecord
    ): void {
        const unitKey = AnalyticsEngine.normalizeText(record.unidadGerencial);
        if (!unitKey) {
            return;
        }

        const bucket = map.get(unitKey) || {
            unitKey,
            label: unitKey,
            filas: 0,
            regiones: new Set<string>(),
            colegios: new Set<string>(),
            solicitudes: new Set<string>(),
            beneficiarios: 0,
            montoIntervencion: 0,
            montoAsignado: 0,
            montoTransferencia: 0,
            montoRetirado: 0,
            solicitudesRevision: new Set<string>(),
            solicitudesCulminadas: new Set<string>(),
            colegiosConTransferencia: new Set<string>(),
            colegiosConRetiro: new Set<string>(),
            grupos: new Map<string, GroupMetricBucket>()
        };

        AnalyticsEngine.addValidDistinct(bucket.solicitudes, record.idSolicitud);
        AnalyticsEngine.addValidDistinct(bucket.regiones, record.region);
        AnalyticsEngine.addValidDistinct(bucket.colegios, record.codigoLocal);
        bucket.beneficiarios += Number(record.beneficiarios) || 0;
        bucket.montoIntervencion += Number(record.montoIntervencion) || 0;
        bucket.montoAsignado += Number(record.montoAsignado) || 0;
        bucket.montoTransferencia += Number(record.montoTransferencia) || 0;
        bucket.montoRetirado += Number(record.montoRetirado) || 0;

        const estado = AnalyticsEngine.normalizeText(record.estadoSolicitud);
        if (AnalyticsEngine.isRevisionState(estado)) {
            AnalyticsEngine.addValidDistinct(bucket.solicitudesRevision, record.idSolicitud);
        }

        if (AnalyticsEngine.isCulminadoState(estado)) {
            AnalyticsEngine.addValidDistinct(bucket.solicitudesCulminadas, record.idSolicitud);
        }

        if (record.montoTransferencia > 0) {
            AnalyticsEngine.addValidDistinct(bucket.colegiosConTransferencia, record.codigoLocal);
        }

        if (record.montoRetirado > 0) {
            AnalyticsEngine.addValidDistinct(bucket.colegiosConRetiro, record.codigoLocal);
        }

        if (unitKey === "UGME") {
            AnalyticsEngine.updateUgmeGroupBucket(bucket, record);
        }

        bucket.filas += 1;
        map.set(unitKey, bucket);
    }

    private static updateUgmeGroupBucket(bucket: UnitKpiBucket, record: SourceRecord): void {
        const grupoKey = AnalyticsEngine.normalizeGrupo(record.grupo);
        if (!grupoKey) {
            return;
        }

        const groupBucket = bucket.grupos.get(grupoKey) || {
            grupoKey,
            grupoLabel: grupoKey,
            montoIntervencion: 0,
            beneficiarios: 0,
            cantidad: 0,
            colegios: new Set<string>(),
            regiones: new Set<string>(),
            filas: 0
        };

        groupBucket.filas += 1;
        AnalyticsEngine.addValidDistinct(groupBucket.colegios, record.codigoLocal);
        AnalyticsEngine.addValidDistinct(groupBucket.regiones, record.region);
        groupBucket.montoIntervencion += Number(record.montoIntervencion) || 0;
        groupBucket.beneficiarios += Number(record.beneficiarios) || 0;
        groupBucket.cantidad += Number(record.cantidad) || 0;
        bucket.grupos.set(grupoKey, groupBucket);
    }

    private static addValidDistinct(set: Set<string>, value: string): void {
        const normalizedValue = AnalyticsEngine.normalizeText(value);
        if (normalizedValue && normalizedValue !== "-" && normalizedValue !== "NULL" && normalizedValue !== "UNDEFINED") {
            set.add(normalizedValue);
        }
    }

    private unitKpiBucketToSummary(bucket: UnitKpiBucket | undefined, unit: UnitKpiName): UnitKpiSummary {
        if (!bucket) {
            return {
                unit,
                unitKey: unit,
                unidad: unit,
                label: unit,
                filas: 0,
                solicitudes: 0,
                regiones: 0,
                colegios: 0,
                montoIntervencion: 0,
                montoAsignado: 0,
                montoTransferencia: 0,
                montoRetirado: 0,
                beneficiarios: 0,
                solicitudesRevision: 0,
                solicitudesCulminadas: 0,
                colegiosConTransferencia: 0,
                colegiosConRetiro: 0,
                grupos: AnalyticsEngine.emptyGroupSummaries(),
                registros: 0
            };
        }

        return {
            unit,
            unitKey: unit,
            unidad: bucket.label,
            label: bucket.label,
            filas: bucket.filas,
            solicitudes: bucket.solicitudes.size,
            regiones: bucket.regiones.size,
            colegios: bucket.colegios.size,
            montoIntervencion: bucket.montoIntervencion,
            montoAsignado: bucket.montoAsignado,
            montoTransferencia: bucket.montoTransferencia,
            montoRetirado: bucket.montoRetirado,
            beneficiarios: bucket.beneficiarios,
            solicitudesRevision: bucket.solicitudesRevision.size,
            solicitudesCulminadas: bucket.solicitudesCulminadas.size,
            colegiosConTransferencia: bucket.colegiosConTransferencia.size,
            colegiosConRetiro: bucket.colegiosConRetiro.size,
            grupos: AnalyticsEngine.groupBucketsToSummaries(bucket.grupos),
            registros: bucket.filas
        };
    }

    private static emptyGroupSummaries(): GroupMetricSummary[] {
        return ugmeGrupoOrder.map((grupo: string) => ({
            grupo,
            montoIntervencion: 0,
            beneficiarios: 0,
            cantidad: 0,
            colegios: 0,
            regiones: 0,
            filas: 0
        }));
    }

    private static groupBucketsToSummaries(groups: Map<string, GroupMetricBucket>): GroupMetricSummary[] {
        return ugmeGrupoOrder.map((grupo: string) => {
            const bucket = groups.get(grupo);
            if (!bucket) {
                return {
                    grupo,
                    montoIntervencion: 0,
                    beneficiarios: 0,
                    cantidad: 0,
                    colegios: 0,
                    regiones: 0,
                    filas: 0
                };
            }

            return {
                grupo: bucket.grupoLabel,
                montoIntervencion: bucket.montoIntervencion,
                beneficiarios: bucket.beneficiarios,
                cantidad: bucket.cantidad,
                colegios: bucket.colegios.size,
                regiones: bucket.regiones.size,
                filas: bucket.filas
            };
        });
    }

    private static addRelation(map: Map<string, Set<string>>, parentKey: string, childKey: string): void {
        let children = map.get(parentKey);
        if (!children) {
            children = new Set<string>();
            map.set(parentKey, children);
        }

        children.add(childKey);
    }

    private bucketToSummary(bucket: MetricBucket, level?: SummaryLevel): BucketSummary {
        const scoreRiesgo = this.getRiskScore(bucket);

        return {
            key: bucket.key,
            label: bucket.label,
            level,
            filas: bucket.filas,
            colegiosUnicos: bucket.colegios.size,
            solicitudesUnicas: bucket.solicitudes.size,
            regionesUnicas: bucket.regiones.size,
            provinciasUnicas: bucket.provincias.size,
            distritosUnicos: bucket.distritos.size,
            unidadesUnicas: bucket.unidades.size,
            montoTotal: bucket.montoTotal,
            beneficiariosTotal: bucket.beneficiariosTotal,
            riesgo: bucket.riesgo,
            riesgoColegios: {
                bajo: bucket.riesgoColegios.bajo.size,
                medio: bucket.riesgoColegios.medio.size,
                alto: bucket.riesgoColegios.alto.size,
                muyAlto: bucket.riesgoColegios.muyAlto.size
            },
            bajo: bucket.riesgo.bajo,
            medio: bucket.riesgo.medio,
            alto: bucket.riesgo.alto,
            muyAlto: bucket.riesgo.muyAlto,
            porcentajeCritico: this.getCriticalPercent(bucket),
            scoreRiesgo,
            clasificacion: this.classifyRisk(scoreRiesgo)
        };
    }

    private static normalizeRiskLevel(value: PrimitiveValue): RiskLevel {
        const normalized = AnalyticsEngine.normalizeKey(value).replace(/\s+/g, " ");

        if (normalized === "muy alto" || normalized === "muyalto") {
            return "Muy Alto";
        }

        if (normalized === "alto") {
            return "Alto";
        }

        if (normalized === "medio") {
            return "Medio";
        }

        return "Bajo";
    }

    private classifyRisk(score: number): RiskLevel {
        if (score < 1.75) {
            return "Bajo";
        }

        if (score < 2.5) {
            return "Medio";
        }

        if (score < 3.25) {
            return "Alto";
        }

        return "Muy Alto";
    }

    private getRiskScore(bucket: MetricBucket): number {
        const totalRiesgo = this.getRiskTotal(bucket);

        if (!totalRiesgo) {
            return 0;
        }

        return (
            bucket.riesgo.bajo +
            bucket.riesgo.medio * 2 +
            bucket.riesgo.alto * 3 +
            bucket.riesgo.muyAlto * 4
        ) / totalRiesgo;
    }

    private getCriticalPercent(bucket: MetricBucket): number {
        const totalRiesgo = this.getRiskTotal(bucket);

        if (!totalRiesgo) {
            return 0;
        }

        return ((bucket.riesgo.alto + bucket.riesgo.muyAlto) / totalRiesgo) * 100;
    }

    private getRiskTotal(bucket: MetricBucket): number {
        return bucket.riesgo.bajo + bucket.riesgo.medio + bucket.riesgo.alto + bucket.riesgo.muyAlto;
    }

    private mapToSummaries(map: Map<string, MetricBucket>, level?: SummaryLevel): BucketSummary[] {
        return Array.from(map.values())
            .map((bucket: MetricBucket) => this.bucketToSummary(bucket, level))
            .sort((left: BucketSummary, right: BucketSummary) => left.label.localeCompare(right.label));
    }

    private relatedSummaries(
        relationMap: Map<string, Set<string>>,
        bucketMap: Map<string, MetricBucket>,
        parentKey: string,
        level: SummaryLevel
    ): BucketSummary[] {
        const relatedKeys = relationMap.get(parentKey) || relationMap.get(AnalyticsEngine.normalizeLookupKey(parentKey));

        if (!relatedKeys) {
            return [];
        }

        return Array.from(relatedKeys)
            .map((key: string) => bucketMap.get(key))
            .filter((bucket: MetricBucket | undefined): bucket is MetricBucket => !!bucket)
            .map((bucket: MetricBucket) => this.bucketToSummary(bucket, level))
            .sort((left: BucketSummary, right: BucketSummary) => left.label.localeCompare(right.label));
    }

    private getBucketSummary(map: Map<string, MetricBucket>, key: string, level: SummaryLevel): BucketSummary | undefined {
        const bucket = map.get(key) || map.get(AnalyticsEngine.normalizeLookupKey(key));
        return bucket ? this.bucketToSummary(bucket, level) : undefined;
    }

    private buildRiskSummaries(): RiskLevelSummary[] {
        const totalColegios = this.result.totals.colegios.size;

        return this.mapToSummaries(this.result.byRiesgo)
            .map((summary: BucketSummary) => ({
                ...summary,
                porcentajeSobreTotal: totalColegios ? (summary.colegiosUnicos / totalColegios) * 100 : 0
            }))
            .sort((left: RiskLevelSummary, right: RiskLevelSummary) => AnalyticsEngine.riskSort(left.label) - AnalyticsEngine.riskSort(right.label));
    }

    private measureLazySampleQuery(): void {
        const start = performance.now();
        const firstRegionKey = Array.from(this.result.byRegion.keys())[0];
        const firstProvinciaKey = firstRegionKey
            ? this.firstRelatedKey(this.result.provinciasByRegion, firstRegionKey)
            : "";
        const firstDistritoKey = firstProvinciaKey
            ? this.firstRelatedKey(this.result.distritosByProvincia, firstProvinciaKey)
            : "";

        if (firstRegionKey) {
            this.getProvincias(firstRegionKey);
        }

        if (firstProvinciaKey) {
            this.getDistritos(firstProvinciaKey);
        }

        if (firstDistritoKey) {
            this.getColegiosByDistrito(firstDistritoKey);
        }

        this.lazySampleQueryMs = performance.now() - start;
    }

    private firstRelatedKey(relationMap: Map<string, Set<string>>, parentKey: string): string {
        const set = relationMap.get(parentKey) || relationMap.get(AnalyticsEngine.normalizeLookupKey(parentKey));
        return set ? Array.from(set)[0] || "" : "";
    }

    private relatedRows(
        relationMap: Map<string, Set<string>>,
        bucketMap: Map<string, MetricBucket>,
        parentKey: string,
        take: number
    ): string[][] {
        const relatedKeys = relationMap.get(parentKey) || relationMap.get(AnalyticsEngine.normalizeLookupKey(parentKey));

        if (!relatedKeys) {
            return [];
        }

        return Array.from(relatedKeys)
            .slice(0, take)
            .map((key: string) => {
                const bucket = bucketMap.get(key);
                return [bucket?.label || key, bucket?.filas.toString() || "0"];
            });
    }

    private static displayValue(value: PrimitiveValue): string {
        if (value === null || value === undefined || value === "") {
            return "(blank)";
        }

        if (value instanceof Date) {
            return value.toISOString();
        }

        return String(value);
    }

    private static normalizeLookupKey(value: string): string {
        return value
            .normalize("NFD")
            .replace(/[\u0300-\u036f]/g, "")
            .trim()
            .toLowerCase();
    }

    private static riskSort(label: string): number {
        if (label === "Bajo") {
            return 1;
        }

        if (label === "Medio") {
            return 2;
        }

        if (label === "Alto") {
            return 3;
        }

        return 4;
    }
}

export class Visual implements IVisual {
    private host: IVisualHost;
    private target: HTMLElement;
    private accumulatedRows: SourceRecord[] = [];
    private accumulatedRowKeys: Set<string> = new Set<string>();
    private lastFetchClickTime?: number;
    private isFetching = false;
    private allDataLoaded = false;
    private accumulatedRowCount = 0;
    private fetchCount = 0;
    private analyticsCache: AnalyticsEngine | null = null;
    private lastDatasetSignature = "";
    private lastIncomingDataSignature = "";
    private segmentLoadStartTime?: number;
    private updateCount = 0;
    private currentView: VisualView = "summary";
    private selectedRiskView: RiskLevel = "Alto";
    private selectedRegion: string | null = null;
    private warnedMissingRegionMatches: Set<string> = new Set<string>();
    private latestDiagnostics: TableDiagnostics | null = null;
    private detailModalUnit: UnitKpiName | null = null;
    private isDetailModalOpen = false;
    private detailSearchText = "";
    private detailPage = 1;
    private detailPageSize = 20;
    private detailCacheByUnit: Map<string, UnitDetailRow[]> = new Map<string, UnitDetailRow[]>();
    private formattingSettings: VisualFormattingSettingsModel = new VisualFormattingSettingsModel();
    private formattingSettingsService: FormattingSettingsService;
    private handleKeydown = (event: KeyboardEvent): void => {
        if (event.key === "Escape" && this.isDetailModalOpen) {
            this.closeDetailModal();
        }
    };

    constructor(options: VisualConstructorOptions) {
        this.host = options.host;
        this.target = options.element;
        this.formattingSettingsService = new FormattingSettingsService();
        this.target.classList.add("data-view-diagnostic");
        window.addEventListener("keydown", this.handleKeydown);
    }

    public destroy(): void {
        window.removeEventListener("keydown", this.handleKeydown);
    }

    public update(options: VisualUpdateOptions): void {
        this.updateCount += 1;
        const updateStart = performance.now();
        const fetchWaitMs = this.lastFetchClickTime
            ? performance.now() - this.lastFetchClickTime
            : undefined;
        this.lastFetchClickTime = undefined;

        const dataViews: DataView[] = options.dataViews || [];
        const dataView = dataViews[0];

        if (dataView) {
            this.formattingSettings = this.formattingSettingsService.populateFormattingSettingsModel(
                VisualFormattingSettingsModel,
                dataView
            );
        }

        const diagnostics = this.buildDiagnostics(dataViews, fetchWaitMs);
        this.latestDiagnostics = diagnostics;

        const renderStart = performance.now();
        this.render(diagnostics);
        diagnostics.performanceMetrics.renderMs = performance.now() - renderStart;
        diagnostics.performanceMetrics.updateTotalMs = performance.now() - updateStart;
        diagnostics.performanceMetrics.rowsPerSecond = this.calculateProcessingRate(
            diagnostics.accumulatedFilas,
            diagnostics.performanceMetrics.updateTotalMs
        );
        diagnostics.performanceMetrics.msPer10kRows = this.calculateMsPer10k(
            diagnostics.accumulatedFilas,
            diagnostics.performanceMetrics.updateTotalMs
        );
        this.updatePerformancePanel(diagnostics);
        this.updatePerformanceMini(diagnostics);
        this.logDiagnostics(diagnostics);
        console.table(diagnostics.performanceMetrics);
    }

    private renderCurrentView(): void {
        if (this.latestDiagnostics) {
            this.render(this.latestDiagnostics);
        }
    }

    public getFormattingModel(): powerbi.visuals.FormattingModel {
        return this.formattingSettingsService.buildFormattingModel(this.formattingSettings);
    }

    private buildDiagnostics(dataViews: DataView[], fetchWaitMs?: number): TableDiagnostics {
        const performanceMetrics: PerformanceMetrics = {
            updateTotalMs: 0,
            segmentLoadMs: 0,
            readDataViewMs: 0,
            findColumnsMs: 0,
            buildRecordsMs: 0,
            accumulateRowsMs: 0,
            buildAnalyticsEngineMs: 0,
            buildFlatIndexesMs: 0,
            buildRelationIndexesMs: 0,
            lazySampleQueryMs: 0,
            renderMs: 0,
            indexedRegiones: 0,
            indexedProvincias: 0,
            indexedDistritos: 0,
            indexedColegios: 0,
            fetchCount: this.fetchCount,
            rowsPerSecond: 0,
            msPer10kRows: 0,
            analyticsCacheHit: false,
            analyticsRebuilt: false,
            datasetSignature: "",
            fetchWaitMs
        };

        const readDataViewResult = measure("readDataView", () => {
            const dataView = dataViews[0];
            const table = dataView?.table;

            return {
                dataView,
                table,
                columns: table?.columns || [],
                rows: table?.rows || [],
                hasMoreData: !!dataView?.metadata?.segment,
                dataReductionText: this.describeDataReduction(table)
            };
        });
        performanceMetrics.readDataViewMs = readDataViewResult.ms;

        const { columns, rows, table, hasMoreData, dataReductionText } = readDataViewResult.value;
        const findColumnsResult = measure("findColumns", () => this.findColumnIndexes(columns));
        performanceMetrics.findColumnsMs = findColumnsResult.ms;

        const columnIndexes = findColumnsResult.value;
        const buildRecordsResult = measure("buildRecords", () => this.buildRecords(rows, columnIndexes));
        performanceMetrics.buildRecordsMs = buildRecordsResult.ms;

        const records = buildRecordsResult.value;
        this.resetAccumulationIfDatasetChanged(records, columns, hasMoreData);

        const accumulateResult = measure("accumulateRows", () => this.accumulateRecords(records));
        performanceMetrics.accumulateRowsMs = accumulateResult.ms;
        this.accumulatedRowCount = this.accumulatedRows.length;
        if (this.isFetching && fetchWaitMs !== undefined) {
            this.isFetching = false;
        }

        if (hasMoreData) {
            this.allDataLoaded = false;
            if (this.segmentLoadStartTime === undefined) {
                this.segmentLoadStartTime = performance.now();
            }

            if (!this.isFetching) {
                this.isFetching = true;
                this.fetchCount += 1;
                this.lastFetchClickTime = performance.now();
                const accepted = this.host.fetchMoreData();
                console.log("fetchMoreData automatic requested", {
                    accepted,
                    fetchCount: this.fetchCount,
                    accumulatedRows: this.accumulatedRowCount
                });
            }

            performanceMetrics.fetchCount = this.fetchCount;
            performanceMetrics.segmentLoadMs = this.segmentLoadStartTime === undefined
                ? 0
                : performance.now() - this.segmentLoadStartTime;

            return {
                dataViewCount: dataViews.length,
                hasTable: !!table,
                columnIndexes,
                columns,
                rows,
                records,
                segmentFilas: records.length,
                accumulatedFilas: this.accumulatedRows.length,
                hasMoreData,
                isLoading: true,
                fetchCount: this.fetchCount,
                datasetSignature: "",
                analyticsCacheHit: false,
                analyticsRebuilt: false,
                dataReductionText,
                analytics: null,
                performanceMetrics
            };
        }

        this.isFetching = false;
        this.allDataLoaded = true;
        const datasetSignature = this.buildDatasetSignature(columns, hasMoreData);
        performanceMetrics.datasetSignature = datasetSignature;
        performanceMetrics.segmentLoadMs = this.segmentLoadStartTime === undefined
            ? 0
            : performance.now() - this.segmentLoadStartTime;

        let analytics = this.analyticsCache;
        let analyticsCacheHit = false;
        let analyticsRebuilt = false;

        if (analytics && datasetSignature === this.lastDatasetSignature) {
            analyticsCacheHit = true;
        } else {
            const analyticsResult = measure("buildAnalyticsEngine", () => AnalyticsEngine.build(this.accumulatedRows));
            performanceMetrics.buildAnalyticsEngineMs = analyticsResult.ms;
            analytics = analyticsResult.value;
            this.analyticsCache = analytics;
            this.lastDatasetSignature = datasetSignature;
            analyticsRebuilt = true;

            const buildMetrics = analytics.getBuildMetrics();
            performanceMetrics.buildFlatIndexesMs = buildMetrics.buildFlatIndexesMs;
            performanceMetrics.buildRelationIndexesMs = buildMetrics.buildRelationIndexesMs;
            performanceMetrics.lazySampleQueryMs = buildMetrics.lazySampleQueryMs;
        }

        const lazyDiagnostics = analytics.getLazyNavigationDiagnostics();
        performanceMetrics.indexedRegiones = lazyDiagnostics.regionesIndexadas;
        performanceMetrics.indexedProvincias = lazyDiagnostics.provinciasIndexadas;
        performanceMetrics.indexedDistritos = lazyDiagnostics.distritosIndexados;
        performanceMetrics.indexedColegios = lazyDiagnostics.colegiosIndexados;
        performanceMetrics.fetchCount = this.fetchCount;
        performanceMetrics.analyticsCacheHit = analyticsCacheHit;
        performanceMetrics.analyticsRebuilt = analyticsRebuilt;
        performanceMetrics.rowsPerSecond = this.calculateProcessingRate(this.accumulatedRows.length, performanceMetrics.buildAnalyticsEngineMs || performanceMetrics.updateTotalMs);
        performanceMetrics.msPer10kRows = this.calculateMsPer10k(this.accumulatedRows.length, performanceMetrics.buildAnalyticsEngineMs || performanceMetrics.updateTotalMs);
        console.table([
            analytics.getUnitKpi("UGSC"),
            analytics.getUnitKpi("UGM")
        ]);
        console.table(analytics.getUgmeKpi().grupos);

        return {
            dataViewCount: dataViews.length,
            hasTable: !!table,
            columnIndexes,
            columns,
            rows,
            records,
            segmentFilas: records.length,
            accumulatedFilas: this.accumulatedRows.length,
            hasMoreData,
            isLoading: false,
            fetchCount: this.fetchCount,
            datasetSignature,
            analyticsCacheHit,
            analyticsRebuilt,
            dataReductionText,
            analytics,
            performanceMetrics
        };
    }

    private findColumnIndexes(columns: DataViewMetadataColumn[]): ColumnIndexes {
        return {
            region: this.findColumnIndexByRole(columns, "region"),
            provincia: this.findColumnIndexByRole(columns, "provincia"),
            codigoLocal: this.findColumnIndexByRole(columns, "codigoLocal"),
            nombreLocal: this.findColumnIndexByRole(columns, "nombreLocal"),
            unidadGerencial: this.findColumnIndexByRole(columns, "unidadGerencial"),
            montoIntervencion: this.findColumnIndexByRole(columns, "montoIntervencion"),
            beneficiarios: this.findColumnIndexByRole(columns, "beneficiarios"),
            montoAsignado: this.findColumnIndexByRole(columns, "montoAsignado"),
            montoTransferencia: this.findColumnIndexByRole(columns, "montoTransferencia"),
            montoRetirado: this.findColumnIndexByRole(columns, "montoRetirado"),
            grupo: this.findColumnIndexByRole(columns, "grupo"),
            cantidad: this.findColumnIndexByRole(columns, "cantidad"),
            distrito: this.findColumnIndexByRole(columns, "distrito"),
            idSolicitud: this.findColumnIndexByRole(columns, "idSolicitud"),
            estadoSolicitud: this.findColumnIndexByRole(columns, "estadoSolicitud"),
            nivelRiesgo: this.findColumnIndexByRole(columns, "nivelRiesgo"),
            montoInversion: this.findColumnIndexByRole(columns, "montoInversion"),
            latitud: this.findColumnIndexByRole(columns, "latitud"),
            longitud: this.findColumnIndexByRole(columns, "longitud")
        };
    }

    private findColumnIndexByRole(columns: DataViewMetadataColumn[], roleName: string): number {
        return columns.findIndex((column: DataViewMetadataColumn) => !!column.roles?.[roleName]);
    }

    private buildRecords(rows: DataViewTableRow[], columnIndexes: ColumnIndexes): SourceRecord[] {
        if (this.hasMissingRole(columnIndexes)) {
            return [];
        }

        return rows.map((row: DataViewTableRow) => ({
            unidadGerencial: this.readString(row, columnIndexes.unidadGerencial),
            region: this.readString(row, columnIndexes.region),
            provincia: this.readString(row, columnIndexes.provincia),
            codigoLocal: this.readString(row, columnIndexes.codigoLocal),
            nombreLocal: this.readString(row, columnIndexes.nombreLocal, "-"),
            montoIntervencion: this.toNumber(row[columnIndexes.montoIntervencion]),
            beneficiarios: this.toNumber(row[columnIndexes.beneficiarios]),
            montoAsignado: this.toNumber(row[columnIndexes.montoAsignado]),
            montoTransferencia: this.toNumber(row[columnIndexes.montoTransferencia]),
            montoRetirado: this.toNumber(row[columnIndexes.montoRetirado]),
            grupo: this.readString(row, columnIndexes.grupo),
            cantidad: this.toNumber(row[columnIndexes.cantidad]),
            distrito: this.readString(row, columnIndexes.distrito),
            idSolicitud: this.readString(row, columnIndexes.idSolicitud),
            estadoSolicitud: this.readString(row, columnIndexes.estadoSolicitud),
            nivelRiesgo: this.readString(row, columnIndexes.nivelRiesgo),
            montoInversion: this.toNumber(row[columnIndexes.montoInversion]) || this.toNumber(row[columnIndexes.montoIntervencion]),
            latitud: this.toNumber(row[columnIndexes.latitud]),
            longitud: this.toNumber(row[columnIndexes.longitud])
        }));
    }

    private accumulateRecords(records: SourceRecord[]): void {
        records.forEach((record: SourceRecord) => {
            const recordKey = this.recordKey(record);
            if (this.accumulatedRowKeys.has(recordKey)) {
                return;
            }

            this.accumulatedRowKeys.add(recordKey);
            this.accumulatedRows.push(record);
        });
    }

    private logDiagnostics(diagnostics: TableDiagnostics): void {
        console.group(`Update #${this.updateCount}`);
        console.log("Rows acumuladas", diagnostics.accumulatedFilas);
        console.log("Hay mas segmentos", diagnostics.hasMoreData);
        console.log("Analytics cache hit", diagnostics.analyticsCacheHit);
        console.log("Analytics rebuilt", diagnostics.analyticsRebuilt);
        console.log("Tiempo Analytics", diagnostics.performanceMetrics.buildAnalyticsEngineMs);
        console.log("Tiempo Render", diagnostics.performanceMetrics.renderMs);
        console.log("Dataset Signature", diagnostics.datasetSignature || "loading");
        console.groupEnd();
    }

    private render(diagnostics: TableDiagnostics): void {
        this.target.replaceChildren();

        if (!diagnostics.hasTable) {
            this.appendMessage("No table DataView received.");
            this.appendPerformanceMini(diagnostics);
            return;
        }

        if (this.hasMissingRole(diagnostics.columnIndexes)) {
            this.appendMessage("Missing one or more required roles for unit KPI cards.");
            this.appendPerformanceMini(diagnostics);
            return;
        }

        if (diagnostics.isLoading) {
            this.appendFinalLoadingPanel(diagnostics);
            this.appendPerformanceMini(diagnostics);
            return;
        }

        if (!diagnostics.analytics) {
            this.appendMessage("Analytics Engine is not available.");
            this.appendPerformanceMini(diagnostics);
            return;
        }

        if (this.currentView === "ugme") {
            this.appendUgmeView(diagnostics.analytics);
        } else {
            this.appendSummaryView(diagnostics.analytics);
        }

        if (this.isDetailModalOpen && this.detailModalUnit) {
            this.appendDetailModal(this.detailModalUnit);
        }

        this.appendPerformanceMini(diagnostics);
    }

    private appendSummaryView(engine: AnalyticsEngine): void {
        const view = document.createElement("section");
        view.className = "dashboard-root";

        const layout = document.createElement("div");
        layout.className = "main-dashboard-layout";

        const mapPanel = document.createElement("section");
        mapPanel.className = "map-panel";
        this.appendPeruRiskMap(mapPanel, engine);
        layout.appendChild(mapPanel);

        const unitsPanel = document.createElement("section");
        unitsPanel.className = "units-panel";

        const title = document.createElement("h1");
        title.textContent = "RESUMEN POR UNIDADES GERENCIALES";
        unitsPanel.appendChild(title);

        const grid = document.createElement("div");
        grid.className = "unit-grid";

        (["UGRD", "UGSC", "UGEO", "UGM"] as UnitKpiName[]).forEach((unit: UnitKpiName) => {
            grid.appendChild(this.createUnitCard(engine.getUnitKpi(unit), false));
        });

        unitsPanel.appendChild(grid);

        const nextButton = document.createElement("button");
        nextButton.className = "floating-next-button";
        nextButton.type = "button";
        nextButton.textContent = ">";
        nextButton.onclick = () => {
            this.currentView = "ugme";
            this.renderCurrentView();
        };
        unitsPanel.appendChild(nextButton);

        layout.appendChild(unitsPanel);
        view.appendChild(layout);

        this.target.appendChild(view);
    }

    private appendPeruRiskMap(container: HTMLElement, engine: AnalyticsEngine): void {
        const title = document.createElement("h2");
        title.className = "map-title";
        title.textContent = "MAPA DE RIESGO POR REGION";
        container.appendChild(title);

        const map = document.createElement("div");
        map.className = "peru-risk-map";

        const tooltip = document.createElement("div");
        tooltip.className = "region-tooltip";
        tooltip.style.display = "none";
        map.appendChild(tooltip);

        const aggregates = this.getRegionAggregates(engine);
        const aggregateMap = this.createRegionAggregateMap(aggregates);
        const svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
        svg.setAttribute("class", "map-svg");
        svg.setAttribute("viewBox", "0 0 920 1220");
        svg.setAttribute("role", "img");
        svg.setAttribute("aria-label", "Mapa de riesgo por region del Peru");

        peruRegionShapes.forEach((shape: MapRegionShape) => {
            const aggregate = this.findRegionAggregate(shape.name, aggregateMap);
            const selectedCount = aggregate ? this.getRiskValue(aggregate.riskBreakdown, this.selectedRiskView) : 0;
            const ratio = aggregate?.totalColegios ? selectedCount / aggregate.totalColegios : Number.NaN;
            const normalizedShape = this.normalizeRegionForMap(shape.name);
            const selected = !!this.selectedRegion && this.normalizeRegionForMap(this.selectedRegion) === normalizedShape;

            if (!aggregate && !this.warnedMissingRegionMatches.has(normalizedShape)) {
                console.warn("Region sin match de agregados para mapa", shape.name);
                this.warnedMissingRegionMatches.add(normalizedShape);
            }

            const path = document.createElementNS("http://www.w3.org/2000/svg", "path");
            path.setAttribute("d", shape.path);
            path.setAttribute("fill", getRiskRatioColor(ratio));
            path.setAttribute("class", [
                "map-region",
                aggregate ? "" : "map-region-no-data",
                selected ? "map-region-selected" : ""
            ].filter(Boolean).join(" "));
            path.setAttribute("data-region", shape.name);

            path.onmouseenter = (event: MouseEvent) => this.showRegionTooltip(tooltip, map, shape.name, aggregate, selectedCount, ratio, event);
            path.onmousemove = (event: MouseEvent) => this.positionRegionTooltip(tooltip, map, event);
            path.onmouseleave = () => {
                tooltip.style.display = "none";
            };
            path.onclick = () => {
                this.selectedRegion = shape.name;
                console.log("onRegionClick", shape.name, aggregate);
                this.renderCurrentView();
            };

            svg.appendChild(path);

            const label = document.createElementNS("http://www.w3.org/2000/svg", "text");
            label.setAttribute("x", shape.labelX.toString());
            label.setAttribute("y", shape.labelY.toString());
            label.setAttribute("class", "map-region-label");
            label.textContent = this.shortRegionLabel(shape.name);
            svg.appendChild(label);
        });

        map.appendChild(svg);
        this.appendMapNationalSummary(map, engine);
        this.appendRiskToggle(map);
        container.appendChild(map);
    }

    private appendMapNationalSummary(map: HTMLElement, engine: AnalyticsEngine): void {
        const totals = engine.getTotals();
        const summary = document.createElement("aside");
        summary.className = "map-summary-card";

        const title = document.createElement("h3");
        title.textContent = "Resumen nacional";
        summary.appendChild(title);

        [
            ["RG", "N° Regiones", this.formatInteger(totals.regionesUnicas)],
            ["CL", "N° Colegios", this.formatInteger(totals.colegiosUnicos)]
        ].forEach(([icon, label, value]: string[]) => {
            const row = document.createElement("div");
            row.className = "map-summary-row";

            const badge = document.createElement("span");
            badge.textContent = icon;
            row.appendChild(badge);

            const content = document.createElement("div");
            const strong = document.createElement("strong");
            strong.textContent = value;
            content.appendChild(strong);

            const small = document.createElement("small");
            small.textContent = label;
            content.appendChild(small);
            row.appendChild(content);
            summary.appendChild(row);
        });

        map.appendChild(summary);
    }

    private appendRiskToggle(map: HTMLElement): void {
        const toggle = document.createElement("div");
        toggle.className = "map-risk-toggle";

        const label = document.createElement("span");
        label.textContent = "Ver por Riesgo:";
        toggle.appendChild(label);

        riskLevels.forEach((risk: RiskLevel) => {
            const button = document.createElement("button");
            button.className = `map-risk-toggle-button${risk === this.selectedRiskView ? " map-risk-toggle-button-active" : ""}`;
            button.type = "button";
            button.textContent = risk;
            button.onclick = () => {
                this.selectedRiskView = risk;
                this.renderCurrentView();
            };
            toggle.appendChild(button);
        });

        map.appendChild(toggle);
    }

    private showRegionTooltip(
        tooltip: HTMLElement,
        map: HTMLElement,
        regionName: string,
        aggregate: RegionAggregate | undefined,
        selectedCount: number,
        ratio: number,
        event: MouseEvent
    ): void {
        tooltip.replaceChildren();

        const title = document.createElement("strong");
        title.textContent = aggregate?.region || regionName;
        tooltip.appendChild(title);

        const rows = aggregate
            ? [
                ["Riesgo seleccionado", this.selectedRiskView],
                [`Colegios ${this.selectedRiskView}`, this.formatInteger(selectedCount)],
                ["Total colegios", this.formatInteger(aggregate.totalColegios)],
                [`% ${this.selectedRiskView}`, this.formatPercent(Number.isFinite(ratio) ? ratio * 100 : 0)],
                ["Score", aggregate.weightedScore.toFixed(2)],
                ["Clasificacion", aggregate.weightedRisk]
            ]
            : [["Sin datos", "No hay agregados para esta region"]];

        rows.forEach(([label, value]: string[]) => {
            const row = document.createElement("div");
            const labelElement = document.createElement("span");
            labelElement.textContent = label;
            row.appendChild(labelElement);

            const valueElement = document.createElement("b");
            valueElement.textContent = value;
            row.appendChild(valueElement);
            tooltip.appendChild(row);
        });

        tooltip.style.display = "block";
        this.positionRegionTooltip(tooltip, map, event);
    }

    private positionRegionTooltip(tooltip: HTMLElement, map: HTMLElement, event: MouseEvent): void {
        const rect = map.getBoundingClientRect();
        const tooltipWidth = tooltip.offsetWidth || 210;
        const tooltipHeight = tooltip.offsetHeight || 150;
        const left = Math.min(Math.max(event.clientX - rect.left + 14, 8), rect.width - tooltipWidth - 8);
        const top = Math.min(Math.max(event.clientY - rect.top + 14, 8), rect.height - tooltipHeight - 8);
        tooltip.style.left = `${left}px`;
        tooltip.style.top = `${top}px`;
    }

    private getRegionAggregates(engine: AnalyticsEngine): RegionAggregate[] {
        return engine.getRegions().map((summary: BucketSummary) => ({
            region: summary.label,
            totalColegios: summary.colegiosUnicos,
            montoInversion: summary.montoTotal,
            estudiantesBeneficiarios: summary.beneficiariosTotal,
            numeroSolicitudes: summary.solicitudesUnicas,
            risk: summary.clasificacion,
            predominantRisk: this.getPredominantRisk(summary.riesgoColegios),
            weightedRisk: summary.clasificacion,
            weightedScore: summary.scoreRiesgo,
            riskBreakdown: summary.riesgoColegios
        }));
    }

    private createRegionAggregateMap(aggregates: RegionAggregate[]): Map<string, RegionAggregate> {
        const map = new Map<string, RegionAggregate>();
        aggregates.forEach((aggregate: RegionAggregate) => {
            map.set(this.normalizeRegionForMap(aggregate.region), aggregate);
        });
        return map;
    }

    private findRegionAggregate(regionName: string, aggregateMap: Map<string, RegionAggregate>): RegionAggregate | undefined {
        const normalized = this.normalizeRegionForMap(regionName);
        const direct = aggregateMap.get(normalized);
        if (direct) {
            return direct;
        }

        const aliases = this.getRegionAliases(normalized);
        for (const alias of aliases) {
            const aggregate = aggregateMap.get(alias);
            if (aggregate) {
                return aggregate;
            }
        }

        return undefined;
    }

    private getRegionAliases(normalizedRegion: string): string[] {
        const aliases: Record<string, string[]> = {
            "LIMA": ["LIMA METROPOLITANA", "LIMA PROVINCIAS"],
            "LIMA METROPOLITANA": ["LIMA", "LIMA METROPOLITANA"],
            "LIMA PROVINCIAS": ["LIMA", "LIMA PROVINCIAS"],
            "CALLAO": ["PROVINCIA CONSTITUCIONAL DEL CALLAO", "CALLAO"]
        };

        return aliases[normalizedRegion] || [];
    }

    private normalizeRegionForMap(value: string): string {
        return this.normalizeText(value)
            .replace(/^REGION\s+/g, "")
            .replace(/\s+REGION$/g, "")
            .replace(/\s+/g, " ");
    }

    private getPredominantRisk(riskBreakdown: RiskBuckets): RiskLevel {
        return riskLevels.reduce((best: RiskLevel, current: RiskLevel) => (
            this.getRiskValue(riskBreakdown, current) > this.getRiskValue(riskBreakdown, best) ? current : best
        ), "Bajo");
    }

    private getRiskValue(riskBreakdown: RiskBuckets, risk: RiskLevel): number {
        if (risk === "Bajo") {
            return riskBreakdown.bajo;
        }

        if (risk === "Medio") {
            return riskBreakdown.medio;
        }

        if (risk === "Alto") {
            return riskBreakdown.alto;
        }

        return riskBreakdown.muyAlto;
    }

    private shortRegionLabel(region: string): string {
        return region
            .replace("LIMA METROPOLITANA", "LIMA MET.")
            .replace("LIMA PROVINCIAS", "LIMA PROV.");
    }

    private appendUgmeView(engine: AnalyticsEngine): void {
        const view = document.createElement("section");
        view.className = "ugme-view";

        const backButton = document.createElement("button");
        backButton.className = "ugme-back-button";
        backButton.type = "button";
        backButton.textContent = "< Volver";
        backButton.onclick = () => {
            this.currentView = "summary";
            this.renderCurrentView();
        };
        view.appendChild(backButton);

        const kpi = engine.getUnitKpi("UGME");
        const card = this.createUnitCard(kpi, true);
        card.classList.add("ugme-card");
        view.appendChild(card);

        this.target.appendChild(view);
    }

    private createUnitCard(kpi: UnitKpiSummary, showUgmeDetail: boolean): HTMLElement {
        const theme = this.getUnitTheme(kpi.unit);
        const card = document.createElement("section");
        card.className = `unit-card unit-card-${kpi.unit.toLowerCase()}`;
        card.style.setProperty("--unit-color", theme.color);
        card.style.setProperty("--unit-bg", theme.background);

        const header = document.createElement("div");
        header.className = "unit-card-header";

        const icon = document.createElement("div");
        icon.className = "unit-card-icon";
        icon.textContent = theme.icon;
        header.appendChild(icon);

        const titleGroup = document.createElement("div");
        titleGroup.className = "unit-card-title-group";

        const title = document.createElement("h2");
        title.className = "unit-card-title";
        title.textContent = kpi.unit;
        titleGroup.appendChild(title);

        const subtitle = document.createElement("p");
        subtitle.className = "unit-card-subtitle";
        subtitle.textContent = theme.name;
        titleGroup.appendChild(subtitle);
        header.appendChild(titleGroup);

        const detailLink = document.createElement("button");
        detailLink.className = "unit-card-detail-link";
        detailLink.type = "button";
        detailLink.textContent = "Ver detalle";
        detailLink.onclick = () => this.openDetailModal(kpi.unit);
        header.appendChild(detailLink);

        card.appendChild(header);

        const grid = document.createElement("div");
        grid.className = "unit-kpi-grid";
        this.getUnitKpiCells(kpi).forEach((item: UnitKpiCell) => grid.appendChild(this.createKpiCell(item)));
        card.appendChild(grid);

        if (showUgmeDetail) {
            this.appendUgmeGroupCards(card, kpi.grupos);
        }

        return card;
    }

    private getUnitTheme(unit: UnitKpiName): UnitTheme {
        const themes: Record<UnitKpiName, UnitTheme> = {
            UGRD: {
                color: "#f97316",
                background: "#fff7ed",
                icon: "R",
                name: "Unidad Gerencial de Reconstruccion y Descentralizacion"
            },
            UGSC: {
                color: "#7c3aed",
                background: "#f5f3ff",
                icon: "S",
                name: "Unidad Gerencial de Supervision y Convenios"
            },
            UGEO: {
                color: "#2563eb",
                background: "#eff6ff",
                icon: "E",
                name: "Unidad Gerencial de Estudios y Obras"
            },
            UGM: {
                color: "#16a34a",
                background: "#f0fdf4",
                icon: "M",
                name: "Unidad Gerencial de Mantenimiento"
            },
            UGME: {
                color: "#0891b2",
                background: "#ecfeff",
                icon: "ME",
                name: "Unidad Gerencial de Mobiliario y Equipamiento"
            }
        };

        return themes[unit];
    }

    private getUnitKpiCells(kpi: UnitKpiSummary): UnitKpiCell[] {
        if (kpi.unit === "UGSC") {
            return [
                { icon: "RG", label: "N° Regiones", value: this.formatInteger(kpi.regiones) },
                { icon: "CL", label: "N° Colegios", value: this.formatInteger(kpi.colegios) },
                { icon: "SO", label: "N° Solicitudes", value: this.formatInteger(kpi.solicitudes) },
                { icon: "EB", label: "Estudiantes Beneficiarios", value: this.formatInteger(kpi.beneficiarios) },
                { icon: "ER", label: "Solicitudes en revision", value: this.formatInteger(kpi.solicitudesRevision) },
                { icon: "CU", label: "Solicitudes culminadas", value: this.formatInteger(kpi.solicitudesCulminadas) },
                { icon: "MI", label: "Monto de Inversion", value: this.formatNumber(kpi.montoIntervencion) }
            ];
        }

        if (kpi.unit === "UGM") {
            const porcentajeTransferencias = kpi.montoAsignado
                ? (kpi.montoTransferencia / kpi.montoAsignado) * 100
                : 0;

            return [
                { icon: "RG", label: "N° Regiones", value: this.formatInteger(kpi.regiones) },
                { icon: "CL", label: "N° Colegios", value: this.formatInteger(kpi.colegios) },
                { icon: "MP", label: "Monto Programado", value: this.formatNumber(kpi.montoAsignado) },
                { icon: "MT", label: "Monto Transferido", value: this.formatNumber(kpi.montoTransferencia) },
                { icon: "%", label: "% Transferencias", value: this.formatPercent(porcentajeTransferencias) },
                { icon: "MR", label: "Monto Retirado", value: this.formatNumber(kpi.montoRetirado) },
                { icon: "CT", label: "N° Colegios con Transferencia", value: this.formatInteger(kpi.colegiosConTransferencia) },
                { icon: "CR", label: "N° Colegios con Retiro", value: this.formatInteger(kpi.colegiosConRetiro) }
            ];
        }

        if (kpi.unit === "UGME") {
            return [
                { icon: "RG", label: "N° Regiones", value: this.formatInteger(kpi.regiones) },
                { icon: "CL", label: "N° Colegios", value: this.formatInteger(kpi.colegios) },
                { icon: "MI", label: "Monto de Inversion", value: this.formatNumber(kpi.montoIntervencion) }
            ];
        }

        return [
            { icon: "RG", label: "N° Regiones", value: this.formatInteger(kpi.regiones) },
            { icon: "CL", label: "N° Colegios", value: this.formatInteger(kpi.colegios) },
            { icon: "MI", label: "Monto de inversion", value: this.formatNumber(kpi.montoIntervencion) },
            { icon: "EB", label: "Estudiantes beneficiarios", value: this.formatInteger(kpi.beneficiarios) }
        ];
    }

    private createKpiCell(item: UnitKpiCell): HTMLElement {
        const cell = document.createElement("div");
        cell.className = "unit-kpi-cell";

        const icon = document.createElement("span");
        icon.className = "unit-kpi-icon";
        icon.textContent = item.icon;
        cell.appendChild(icon);

        const content = document.createElement("div");

        const value = document.createElement("div");
        value.className = "unit-kpi-value";
        value.textContent = item.value;
        content.appendChild(value);

        const label = document.createElement("div");
        label.className = "unit-kpi-label";
        label.textContent = item.label;
        content.appendChild(label);

        cell.appendChild(content);
        return cell;
    }

    private appendUgmeGroupCards(card: HTMLElement, grupos: GroupMetricSummary[]): void {
        const title = document.createElement("h3");
        title.className = "ugme-section-title";
        title.textContent = "Detalle por grupo";
        card.appendChild(title);

        const grid = document.createElement("div");
        grid.className = "ugme-group-grid";

        grupos.forEach((grupo: GroupMetricSummary) => {
            const groupCard = document.createElement("section");
            groupCard.className = "ugme-group-card";

            const groupTitle = document.createElement("h4");
            groupTitle.textContent = grupo.grupo;
            groupCard.appendChild(groupTitle);

            const rows = [
                ["Cantidad", this.formatInteger(grupo.cantidad)],
                ["Beneficiarios", this.formatInteger(grupo.beneficiarios)],
                ["Monto de Inversion", this.formatNumber(grupo.montoIntervencion)]
            ];

            rows.forEach(([labelText, valueText]: string[]) => {
                const row = document.createElement("div");
                row.className = "ugme-group-row";

                const label = document.createElement("span");
                label.textContent = labelText;
                row.appendChild(label);

                const value = document.createElement("strong");
                value.textContent = valueText;
                row.appendChild(value);

                groupCard.appendChild(row);
            });

            grid.appendChild(groupCard);
        });

        card.appendChild(grid);
    }

    private openDetailModal(unit: UnitKpiName): void {
        this.detailModalUnit = unit;
        this.isDetailModalOpen = true;
        this.detailSearchText = "";
        this.detailPage = 1;
        this.renderCurrentView();
    }

    private closeDetailModal(): void {
        this.isDetailModalOpen = false;
        this.detailModalUnit = null;
        this.detailSearchText = "";
        this.detailPage = 1;
        this.renderCurrentView();
    }

    private appendDetailModal(unit: UnitKpiName): void {
        const backdrop = document.createElement("div");
        backdrop.className = "detail-modal-backdrop";
        backdrop.onclick = (event: MouseEvent) => {
            if (event.target === backdrop) {
                this.closeDetailModal();
            }
        };

        const modal = document.createElement("section");
        modal.className = "detail-modal";

        const header = document.createElement("header");
        header.className = "detail-modal-header";

        const titleGroup = document.createElement("div");
        const title = document.createElement("h2");
        title.textContent = `Detalle de colegios - ${unit}`;
        titleGroup.appendChild(title);

        const subtitle = document.createElement("p");
        subtitle.textContent = "Registros considerados para la unidad seleccionada";
        titleGroup.appendChild(subtitle);
        header.appendChild(titleGroup);

        const closeIcon = document.createElement("button");
        closeIcon.className = "detail-modal-close";
        closeIcon.type = "button";
        closeIcon.textContent = "X";
        closeIcon.onclick = () => this.closeDetailModal();
        header.appendChild(closeIcon);
        modal.appendChild(header);

        const allRows = this.getUnitSchoolDetails(unit);
        const filteredRows = this.filterDetailRows(allRows);
        const totalPages = Math.max(1, Math.ceil(filteredRows.length / this.detailPageSize));
        this.detailPage = Math.min(this.detailPage, totalPages);
        const pageStart = (this.detailPage - 1) * this.detailPageSize;
        const pageRows = filteredRows.slice(pageStart, pageStart + this.detailPageSize);

        const toolbar = document.createElement("div");
        toolbar.className = "detail-modal-toolbar";

        const search = document.createElement("input");
        search.className = "detail-search-input";
        search.type = "search";
        search.placeholder = "Buscar colegio, region, provincia, distrito o grupo";
        search.value = this.detailSearchText;
        search.oninput = () => {
            this.detailSearchText = search.value;
            this.detailPage = 1;
            this.renderCurrentView();
        };
        toolbar.appendChild(search);

        const count = document.createElement("span");
        count.className = "detail-count";
        count.textContent = `Mostrando ${pageRows.length} de ${filteredRows.length} colegios`;
        toolbar.appendChild(count);
        modal.appendChild(toolbar);

        this.appendDetailTable(modal, unit, pageRows);
        this.appendDetailPagination(modal, totalPages);

        const footer = document.createElement("footer");
        footer.className = "detail-modal-footer";

        const closeButton = document.createElement("button");
        closeButton.type = "button";
        closeButton.textContent = "Cerrar";
        closeButton.onclick = () => this.closeDetailModal();
        footer.appendChild(closeButton);
        modal.appendChild(footer);

        backdrop.appendChild(modal);
        this.target.appendChild(backdrop);
    }

    private appendDetailTable(modal: HTMLElement, unit: UnitKpiName, rows: UnitDetailRow[]): void {
        const scroller = document.createElement("div");
        scroller.className = "detail-table-scroller";

        const table = document.createElement("table");
        table.className = "detail-table";
        const thead = document.createElement("thead");
        const headerRow = document.createElement("tr");
        const columns = this.getDetailColumns(unit);

        columns.forEach((column: DetailColumn) => {
            const th = document.createElement("th");
            th.textContent = column.label;
            headerRow.appendChild(th);
        });

        thead.appendChild(headerRow);
        table.appendChild(thead);

        const tbody = document.createElement("tbody");
        if (!rows.length) {
            const row = document.createElement("tr");
            const cell = document.createElement("td");
            cell.colSpan = columns.length;
            cell.textContent = "Sin colegios para mostrar.";
            row.appendChild(cell);
            tbody.appendChild(row);
        } else {
            rows.forEach((detailRow: UnitDetailRow) => {
                const row = document.createElement("tr");
                columns.forEach((column: DetailColumn) => {
                    const td = document.createElement("td");
                    td.textContent = column.value(detailRow);
                    row.appendChild(td);
                });
                tbody.appendChild(row);
            });
        }

        table.appendChild(tbody);
        scroller.appendChild(table);
        modal.appendChild(scroller);
    }

    private appendDetailPagination(modal: HTMLElement, totalPages: number): void {
        const pagination = document.createElement("div");
        pagination.className = "detail-pagination";

        const previous = document.createElement("button");
        previous.type = "button";
        previous.textContent = "Anterior";
        previous.disabled = this.detailPage <= 1;
        previous.onclick = () => {
            this.detailPage = Math.max(1, this.detailPage - 1);
            this.renderCurrentView();
        };
        pagination.appendChild(previous);

        const page = document.createElement("span");
        page.textContent = `Pagina ${this.detailPage} de ${totalPages}`;
        pagination.appendChild(page);

        const next = document.createElement("button");
        next.type = "button";
        next.textContent = "Siguiente";
        next.disabled = this.detailPage >= totalPages;
        next.onclick = () => {
            this.detailPage = Math.min(totalPages, this.detailPage + 1);
            this.renderCurrentView();
        };
        pagination.appendChild(next);

        modal.appendChild(pagination);
    }

    private getUnitSchoolDetails(unit: UnitKpiName): UnitDetailRow[] {
        const cachedRows = this.detailCacheByUnit.get(unit);
        if (cachedRows) {
            return cachedRows;
        }

        const detailMap = new Map<string, UnitDetailRow>();

        this.accumulatedRows.forEach((record: SourceRecord) => {
            const recordUnit = this.normalizeText(record.unidadGerencial) as UnitKpiName;
            if (recordUnit !== unit) {
                return;
            }

            const codigoKey = this.normalizeText(record.codigoLocal);
            const grupoKey = unit === "UGME" ? this.normalizeText(record.grupo) : "";
            const key = unit === "UGME" ? `${codigoKey}|${grupoKey}` : codigoKey;
            const row = detailMap.get(key) || {
                unit,
                codigoLocal: record.codigoLocal || "-",
                nombreLocal: record.nombreLocal || "-",
                region: record.region || "-",
                provincia: record.provincia || "-",
                distrito: record.distrito || "-",
                grupo: unit === "UGME" ? record.grupo || "-" : undefined,
                montoIntervencion: 0,
                beneficiarios: 0,
                cantidad: 0,
                solicitudes: new Set<string>(),
                solicitudesRevision: new Set<string>(),
                solicitudesCulminadas: new Set<string>(),
                montoAsignado: 0,
                montoTransferencia: 0,
                montoRetirado: 0
            };

            row.montoIntervencion += Number(record.montoIntervencion) || 0;
            row.beneficiarios += Number(record.beneficiarios) || 0;
            row.cantidad += Number(record.cantidad) || 0;
            row.montoAsignado += Number(record.montoAsignado) || 0;
            row.montoTransferencia += Number(record.montoTransferencia) || 0;
            row.montoRetirado += Number(record.montoRetirado) || 0;

            this.addValidDetailDistinct(row.solicitudes, record.idSolicitud);
            const estado = this.normalizeText(record.estadoSolicitud);
            if (estado === "EN REVISION" || estado === "REVISION") {
                this.addValidDetailDistinct(row.solicitudesRevision, record.idSolicitud);
            }

            if (estado === "CULMINADO" || estado === "CULMINADA" || estado === "CULMINADOS" || estado === "CULMINADAS") {
                this.addValidDetailDistinct(row.solicitudesCulminadas, record.idSolicitud);
            }

            detailMap.set(key, row);
        });

        const rows = Array.from(detailMap.values())
            .sort((left: UnitDetailRow, right: UnitDetailRow) => {
                const byRegion = left.region.localeCompare(right.region);
                if (byRegion) {
                    return byRegion;
                }

                return left.codigoLocal.localeCompare(right.codigoLocal);
            });

        this.detailCacheByUnit.set(unit, rows);
        return rows;
    }

    private filterDetailRows(rows: UnitDetailRow[]): UnitDetailRow[] {
        const searchText = this.normalizeText(this.detailSearchText);
        if (!searchText) {
            return rows;
        }

        return rows.filter((row: UnitDetailRow) => [
            row.codigoLocal,
            row.nombreLocal,
            row.region,
            row.provincia,
            row.distrito,
            row.grupo || ""
        ].some((value: string) => this.normalizeText(value).includes(searchText)));
    }

    private getDetailColumns(unit: UnitKpiName): DetailColumn[] {
        const baseColumns: DetailColumn[] = [
            { label: "CodigoLocal", value: (row: UnitDetailRow) => row.codigoLocal },
            { label: "NombreLocal", value: (row: UnitDetailRow) => row.nombreLocal || "-" },
            { label: "Region", value: (row: UnitDetailRow) => row.region },
            { label: "Provincia", value: (row: UnitDetailRow) => row.provincia },
            { label: "Distrito", value: (row: UnitDetailRow) => row.distrito }
        ];

        if (unit === "UGSC") {
            return [
                ...baseColumns,
                { label: "N° Solicitudes", value: (row: UnitDetailRow) => this.formatInteger(row.solicitudes.size) },
                { label: "Estudiantes beneficiarios", value: (row: UnitDetailRow) => this.formatInteger(row.beneficiarios) },
                { label: "Solicitudes en revision", value: (row: UnitDetailRow) => this.formatInteger(row.solicitudesRevision.size) },
                { label: "Solicitudes culminadas", value: (row: UnitDetailRow) => this.formatInteger(row.solicitudesCulminadas.size) },
                { label: "Monto de inversion", value: (row: UnitDetailRow) => this.formatCurrencyNoDecimals(row.montoIntervencion) }
            ];
        }

        if (unit === "UGM") {
            return [
                ...baseColumns,
                { label: "Monto Programado", value: (row: UnitDetailRow) => this.formatCurrencyNoDecimals(row.montoAsignado) },
                { label: "Monto Transferido", value: (row: UnitDetailRow) => this.formatCurrencyNoDecimals(row.montoTransferencia) },
                { label: "% Transferencias", value: (row: UnitDetailRow) => this.formatPercent(row.montoAsignado ? (row.montoTransferencia / row.montoAsignado) * 100 : 0) },
                { label: "Monto Retirado", value: (row: UnitDetailRow) => this.formatCurrencyNoDecimals(row.montoRetirado) },
                { label: "Tiene Transferencia", value: (row: UnitDetailRow) => row.montoTransferencia > 0 ? "Sí" : "No" },
                { label: "Tiene Retiro", value: (row: UnitDetailRow) => row.montoRetirado > 0 ? "Sí" : "No" }
            ];
        }

        if (unit === "UGME") {
            return [
                ...baseColumns,
                { label: "Grupo", value: (row: UnitDetailRow) => row.grupo || "-" },
                { label: "Cantidad", value: (row: UnitDetailRow) => this.formatInteger(row.cantidad) },
                { label: "Beneficiarios", value: (row: UnitDetailRow) => this.formatInteger(row.beneficiarios) },
                { label: "Monto de inversion", value: (row: UnitDetailRow) => this.formatCurrencyNoDecimals(row.montoIntervencion) }
            ];
        }

        return [
            ...baseColumns,
            { label: "Monto de inversion", value: (row: UnitDetailRow) => this.formatCurrencyNoDecimals(row.montoIntervencion) },
            { label: "Estudiantes beneficiarios", value: (row: UnitDetailRow) => this.formatInteger(row.beneficiarios) }
        ];
    }

    private addValidDetailDistinct(set: Set<string>, value: string): void {
        const normalizedValue = this.normalizeText(value);
        if (normalizedValue && normalizedValue !== "-" && normalizedValue !== "NULL" && normalizedValue !== "UNDEFINED") {
            set.add(normalizedValue);
        }
    }

    private appendPerformanceMini(diagnostics: TableDiagnostics): void {
        const performance = document.createElement("div");
        performance.className = "performance-mini";
        performance.textContent = `Tiempo total: ${this.formatMs(diagnostics.performanceMetrics.updateTotalMs)}`;
        this.target.appendChild(performance);
    }

    private updatePerformanceMini(diagnostics: TableDiagnostics): void {
        const performance = this.target.querySelector(".performance-mini");
        if (performance) {
            performance.textContent = `Tiempo total: ${this.formatMs(diagnostics.performanceMetrics.updateTotalMs)}`;
        }
    }

    private appendFinalLoadingPanel(diagnostics: TableDiagnostics): void {
        const panel = document.createElement("section");
        panel.className = "unit-summary-view";

        const title = document.createElement("h1");
        title.textContent = "RESUMEN POR UNIDADES GERENCIALES";
        panel.appendChild(title);

        const message = document.createElement("p");
        message.className = "loading-message";
        message.textContent = `Cargando datos acumulados: ${this.formatInteger(diagnostics.accumulatedFilas)} filas`;
        panel.appendChild(message);

        this.target.appendChild(panel);
    }

    private appendLoadingPanel(diagnostics: TableDiagnostics): void {
        const title = document.createElement("h2");
        title.textContent = "Cargando informacion...";
        this.target.appendChild(title);

        this.appendKeyValueList([
            ["Segmentos descargados", diagnostics.fetchCount.toString()],
            ["Filas acumuladas", diagnostics.accumulatedFilas.toString()],
            ["Estado", "Procesando..."],
            ["Mensaje", "Espere un momento."]
        ]);
    }

    private appendDataViewPanel(diagnostics: TableDiagnostics): void {
        this.appendKeyValueList([
            ["options.dataViews.length", diagnostics.dataViewCount.toString()],
            ["table DataView", diagnostics.hasTable ? "si" : "no"],
            ["filas del segmento actual", diagnostics.segmentFilas.toString()],
            ["filas acumuladas", diagnostics.accumulatedFilas.toString()],
            ["hay mas datos / segment", diagnostics.hasMoreData ? "si" : "no"],
            ["dataReduction", diagnostics.dataReductionText]
        ]);
    }

    private appendUnitKpiCards(engine: AnalyticsEngine): void {
        const title = document.createElement("h2");
        title.textContent = "Tarjetas por unidad gerencial";
        this.target.appendChild(title);

        const container = document.createElement("div");
        container.className = "unit-kpi-grid";

        engine.getUnitKpis().forEach((kpi: UnitKpiSummary) => {
            const card = document.createElement("section");
            card.className = "unit-kpi-card";

            const cardTitle = document.createElement("h3");
            cardTitle.textContent = kpi.unit;
            card.appendChild(cardTitle);

            if (kpi.unit === "UGSC") {
                const subtitle = document.createElement("p");
                subtitle.className = "unit-kpi-subtitle";
                subtitle.textContent = "Unidad Gerencial de Supervisión y Convenios";
                card.appendChild(subtitle);
            }

            const list = document.createElement("dl");
            list.className = "unit-kpi-list";

            const items = this.buildUnitKpiItems(kpi);

            items.forEach(([label, value]: string[]) => {
                const term = document.createElement("dt");
                term.textContent = label;
                list.appendChild(term);

                const description = document.createElement("dd");
                description.textContent = value;
                list.appendChild(description);
            });

            card.appendChild(list);

            if (kpi.unit === "UGME") {
                this.appendUgmeGroupDetail(card, kpi.grupos);
            }

            container.appendChild(card);
        });

        this.target.appendChild(container);
    }

    private buildUnitKpiItems(kpi: UnitKpiSummary): string[][] {
        if (kpi.unit === "UGSC") {
            return [
                ["N° Regiones", this.formatInteger(kpi.regiones)],
                ["N° Colegios", this.formatInteger(kpi.colegios)],
                ["N° Solicitudes", this.formatInteger(kpi.solicitudes)],
                ["Estudiantes Beneficiarios", this.formatInteger(kpi.beneficiarios)],
                ["Solicitudes en revisión", this.formatInteger(kpi.solicitudesRevision)],
                ["Solicitudes culminadas", this.formatInteger(kpi.solicitudesCulminadas)],
                ["Monto de Inversión", this.formatNumber(kpi.montoIntervencion)]
            ];
        }

        if (kpi.unit === "UGM") {
            const porcentajeTransferencias = kpi.montoAsignado
                ? (kpi.montoTransferencia / kpi.montoAsignado) * 100
                : 0;

            return [
                ["N° Regiones", this.formatInteger(kpi.regiones)],
                ["N° Colegios", this.formatInteger(kpi.colegios)],
                ["Monto Programado", this.formatNumber(kpi.montoAsignado)],
                ["Monto Transferido", this.formatNumber(kpi.montoTransferencia)],
                ["% Transferencias", this.formatPercent(porcentajeTransferencias)],
                ["Monto Retirado", this.formatNumber(kpi.montoRetirado)],
                ["N° Colegios con Transferencia", this.formatInteger(kpi.colegiosConTransferencia)],
                ["N° Colegios con Retiro", this.formatInteger(kpi.colegiosConRetiro)]
            ];
        }

        if (kpi.unit === "UGME") {
            return [
                ["N° Regiones", this.formatInteger(kpi.regiones)],
                ["N° Colegios", this.formatInteger(kpi.colegios)],
                ["Monto de Inversión", this.formatNumber(kpi.montoIntervencion)]
            ];
        }

        return [
            ["Nro. Regiones", this.formatInteger(kpi.regiones)],
            ["Nro. Colegios", this.formatInteger(kpi.colegios)],
            ["Monto de inversion", this.formatNumber(kpi.montoIntervencion)],
            ["Estudiantes beneficiarios", this.formatInteger(kpi.beneficiarios)]
        ];
    }

    private appendUgmeGroupDetail(card: HTMLElement, grupos: GroupMetricSummary[]): void {
        const subtitle = document.createElement("h4");
        subtitle.textContent = "Detalle por grupo";
        card.appendChild(subtitle);

        const table = this.createTable([
            "Grupo",
            "Cantidad",
            "Beneficiarios",
            "Monto de Inversión"
        ]);
        const tbody = document.createElement("tbody");

        grupos.forEach((grupo: GroupMetricSummary) => {
            this.appendTableRow(tbody, [
                grupo.grupo,
                this.formatInteger(grupo.cantidad),
                this.formatInteger(grupo.beneficiarios),
                this.formatNumber(grupo.montoIntervencion)
            ]);
        });

        table.appendChild(tbody);
        card.appendChild(table);
    }

    private appendGlobalTotalsPanel(engine: AnalyticsEngine): void {
        const title = document.createElement("h2");
        title.textContent = "Totales globales";
        this.target.appendChild(title);

        const totals = engine.getTotals();

        this.appendKeyValueList([
            ["filas", totals.filas.toString()],
            ["colegios unicos", totals.colegiosUnicos.toString()],
            ["solicitudes unicas", totals.solicitudesUnicas.toString()],
            ["regiones", totals.regionesUnicas.toString()],
            ["provincias", totals.provinciasUnicas.toString()],
            ["distritos", totals.distritosUnicos.toString()],
            ["unidades gerenciales", totals.unidadesUnicas.toString()],
            ["monto total", this.formatNumber(totals.montoTotal)],
            ["% critico global", this.formatPercent(totals.porcentajeCritico)],
            ["score global", totals.scoreRiesgo.toFixed(2)],
            ["clasificacion global", totals.clasificacion]
        ]);
    }

    private appendLazyNavigationDiagnosticsPanel(engine: AnalyticsEngine): void {
        const title = document.createElement("h2");
        title.textContent = "Diagnostico de navegacion lazy";
        this.target.appendChild(title);

        const diagnostics = engine.getLazyNavigationDiagnostics();

        this.appendKeyValueList([
            ["regiones indexadas", diagnostics.regionesIndexadas.toString()],
            ["provincias indexadas", diagnostics.provinciasIndexadas.toString()],
            ["distritos indexados", diagnostics.distritosIndexados.toString()],
            ["colegios indexados", diagnostics.colegiosIndexados.toString()],
            ["relaciones provinciasByRegion", diagnostics.relacionesProvinciasByRegion.toString()],
            ["relaciones distritosByProvincia", diagnostics.relacionesDistritosByProvincia.toString()],
            ["relaciones colegiosByDistrito", diagnostics.relacionesColegiosByDistrito.toString()],
            ["primera region", diagnostics.primeraRegion],
            ["provincias asociadas", diagnostics.provinciasPrimeraRegion.toString()],
            ["primera provincia", diagnostics.primeraProvincia],
            ["distritos asociados", diagnostics.distritosPrimeraProvincia.toString()],
            ["primer distrito", diagnostics.primerDistrito],
            ["colegios asociados", diagnostics.colegiosPrimerDistrito.toString()],
            ["nota", diagnostics.notaProvinciaDistrito]
        ]);

        this.appendDiagnosticRowsTable(
            "Primeras 5 provincias de la primera region",
            ["Provincia", "Filas"],
            diagnostics.primerasProvincias
        );
        this.appendDiagnosticRowsTable(
            "Primeros 5 distritos de la primera provincia",
            ["Distrito", "Filas"],
            diagnostics.primerosDistritos
        );
        this.appendDiagnosticRowsTable(
            "Primeros 5 colegios del primer distrito",
            ["Colegio", "Filas"],
            diagnostics.primerosColegios
        );
    }

    private appendPerformancePanel(diagnostics: TableDiagnostics): void {
        const title = document.createElement("h2");
        title.textContent = "Performance";
        this.target.appendChild(title);

        const metrics = this.getPerformanceDisplayRows(diagnostics);
        const list = document.createElement("dl");
        list.className = "metric-list performance-list";

        metrics.forEach(([key, label, value]: string[]) => {
            const term = document.createElement("dt");
            term.textContent = label;
            list.appendChild(term);

            const description = document.createElement("dd");
            description.dataset.perfKey = key;
            description.textContent = value;
            list.appendChild(description);
        });

        this.target.appendChild(list);
    }

    private updatePerformancePanel(diagnostics: TableDiagnostics): void {
        this.getPerformanceDisplayRows(diagnostics).forEach(([key, , value]: string[]) => {
            const element = this.target.querySelector(`[data-perf-key="${key}"]`);
            if (element) {
                element.textContent = value;
            }
        });
    }

    private getPerformanceDisplayRows(diagnostics: TableDiagnostics): string[][] {
        const metrics = diagnostics.performanceMetrics;

        return [
            ["updateTotalMs", "Update total", this.formatMs(metrics.updateTotalMs)],
            ["segmentLoadMs", "Tiempo carga segmentos", this.formatMs(metrics.segmentLoadMs)],
            ["readDataViewMs", "Read DataView", this.formatMs(metrics.readDataViewMs)],
            ["findColumnsMs", "Find columns", this.formatMs(metrics.findColumnsMs)],
            ["buildRecordsMs", "Build records", this.formatMs(metrics.buildRecordsMs)],
            ["accumulateRowsMs", "Accumulate rows", this.formatMs(metrics.accumulateRowsMs)],
            ["buildAnalyticsEngineMs", "Build analytics engine", this.formatMs(metrics.buildAnalyticsEngineMs)],
            ["buildFlatIndexesMs", "Build flat indexes", this.formatMs(metrics.buildFlatIndexesMs)],
            ["buildRelationIndexesMs", "Build relation indexes", this.formatMs(metrics.buildRelationIndexesMs)],
            ["lazySampleQueryMs", "Lazy sample query", this.formatMs(metrics.lazySampleQueryMs)],
            ["renderMs", "Render HTML", this.formatMs(metrics.renderMs)],
            ["fetchWaitMs", "Fetch wait", metrics.fetchWaitMs === undefined ? "n/a" : this.formatMs(metrics.fetchWaitMs)],
            ["indexedRegiones", "Regiones indexadas", metrics.indexedRegiones.toString()],
            ["indexedProvincias", "Provincias indexadas", metrics.indexedProvincias.toString()],
            ["indexedDistritos", "Distritos indexados", metrics.indexedDistritos.toString()],
            ["indexedColegios", "Colegios indexados", metrics.indexedColegios.toString()],
            ["fetchCount", "Segmentos recibidos", metrics.fetchCount.toString()],
            ["analyticsCacheHit", "Analytics Cache", metrics.analyticsCacheHit ? "SI" : "NO"],
            ["analyticsRebuilt", "Analytics Rebuilt", metrics.analyticsRebuilt ? "SI" : "NO"],
            ["datasetSignature", "Dataset Signature", metrics.datasetSignature || "n/a"],
            ["segmentFilas", "Rows segment", diagnostics.segmentFilas.toString()],
            ["accumulatedFilas", "Rows accumulated", diagnostics.accumulatedFilas.toString()],
            ["processingRate", "Rows/second", `${metrics.rowsPerSecond.toFixed(2)} rows/s`],
            ["msPer10k", "Ms por 10,000 filas", this.formatMs(metrics.msPer10kRows)]
        ];
    }

    private appendRegionSummaryTable(engine: AnalyticsEngine): void {
        const title = document.createElement("h2");
        title.textContent = "Resumen por region";
        this.target.appendChild(title);

        const table = this.createTable([
            "Region",
            "Filas",
            "Colegios unicos",
            "Solicitudes unicas",
            "Monto total",
            "Bajo",
            "Medio",
            "Alto",
            "Muy Alto",
            "% critico",
            "Score riesgo",
            "Clasificacion"
        ]);
        const tbody = document.createElement("tbody");
        const summaries = engine.getRegions();

        summaries.forEach((summary: BucketSummary) => {
            this.appendTableRow(tbody, [
                summary.label,
                summary.filas.toString(),
                summary.colegiosUnicos.toString(),
                summary.solicitudesUnicas.toString(),
                this.formatNumber(summary.montoTotal),
                summary.riesgo.bajo.toString(),
                summary.riesgo.medio.toString(),
                summary.riesgo.alto.toString(),
                summary.riesgo.muyAlto.toString(),
                this.formatPercent(summary.porcentajeCritico),
                summary.scoreRiesgo.toFixed(2),
                summary.clasificacion
            ]);
        });

        const totals = engine.getTotals();
        this.appendTableRow(tbody, [
            "TOTAL",
            totals.filas.toString(),
            totals.colegiosUnicos.toString(),
            totals.solicitudesUnicas.toString(),
            this.formatNumber(totals.montoTotal),
            totals.riesgo.bajo.toString(),
            totals.riesgo.medio.toString(),
            totals.riesgo.alto.toString(),
            totals.riesgo.muyAlto.toString(),
            this.formatPercent(totals.porcentajeCritico),
            totals.scoreRiesgo.toFixed(2),
            totals.clasificacion
        ], "summary-total-row");

        table.appendChild(tbody);
        this.target.appendChild(table);
    }

    private appendUnitSummaryTable(engine: AnalyticsEngine): void {
        const title = document.createElement("h2");
        title.textContent = "Por unidad gerencial";
        this.target.appendChild(title);

        const table = this.createTable([
            "Unidad",
            "Filas",
            "Colegios unicos",
            "Solicitudes unicas",
            "Monto total",
            "% critico",
            "Clasificacion"
        ]);
        const tbody = document.createElement("tbody");

        engine.getUnidades().forEach((summary: BucketSummary) => {
            this.appendTableRow(tbody, [
                summary.label,
                summary.filas.toString(),
                summary.colegiosUnicos.toString(),
                summary.solicitudesUnicas.toString(),
                this.formatNumber(summary.montoTotal),
                this.formatPercent(summary.porcentajeCritico),
                summary.clasificacion
            ]);
        });

        table.appendChild(tbody);
        this.target.appendChild(table);
    }

    private appendStateSummaryTable(engine: AnalyticsEngine): void {
        const title = document.createElement("h2");
        title.textContent = "Por estado";
        this.target.appendChild(title);

        const table = this.createTable(["Estado", "Filas", "Colegios unicos", "Solicitudes unicas", "Monto total"]);
        const tbody = document.createElement("tbody");

        engine.getEstados().forEach((summary: BucketSummary) => {
            this.appendTableRow(tbody, [
                summary.label,
                summary.filas.toString(),
                summary.colegiosUnicos.toString(),
                summary.solicitudesUnicas.toString(),
                this.formatNumber(summary.montoTotal)
            ]);
        });

        table.appendChild(tbody);
        this.target.appendChild(table);
    }

    private appendRiskSummaryTable(engine: AnalyticsEngine): void {
        const title = document.createElement("h2");
        title.textContent = "Por nivel de riesgo";
        this.target.appendChild(title);

        const table = this.createTable(["Nivel", "Filas", "Colegios unicos", "Solicitudes unicas", "Monto total", "% sobre total de colegios"]);
        const tbody = document.createElement("tbody");

        engine.getRiesgos().forEach((summary: RiskLevelSummary) => {
            this.appendTableRow(tbody, [
                summary.label,
                summary.filas.toString(),
                summary.colegiosUnicos.toString(),
                summary.solicitudesUnicas.toString(),
                this.formatNumber(summary.montoTotal),
                this.formatPercent(summary.porcentajeSobreTotal)
            ]);
        });

        table.appendChild(tbody);
        this.target.appendChild(table);
    }

    private appendColumnsTable(columns: DataViewMetadataColumn[]): void {
        const title = document.createElement("h2");
        title.textContent = "Columnas";
        this.target.appendChild(title);

        const table = this.createTable(["Index", "Column", "Roles"]);
        const tbody = document.createElement("tbody");

        columns.forEach((column: DataViewMetadataColumn, index: number) => {
            this.appendTableRow(tbody, [
                index.toString(),
                column.displayName || "(sin nombre)",
                JSON.stringify(column.roles || {})
            ]);
        });

        table.appendChild(tbody);
        this.target.appendChild(table);
    }

    private appendDiagnosticRowsTable(titleText: string, headers: string[], rows: string[][]): void {
        const title = document.createElement("h3");
        title.textContent = titleText;
        this.target.appendChild(title);

        const table = this.createTable(headers);
        const tbody = document.createElement("tbody");

        if (!rows.length) {
            this.appendTableRow(tbody, ["(sin datos)", "0"]);
        } else {
            rows.forEach((row: string[]) => this.appendTableRow(tbody, row));
        }

        table.appendChild(tbody);
        this.target.appendChild(table);
    }

    private appendKeyValueList(metrics: string[][]): void {
        const list = document.createElement("dl");
        list.className = "metric-list";

        metrics.forEach(([label, value]: string[]) => {
            const term = document.createElement("dt");
            term.textContent = label;
            list.appendChild(term);

            const description = document.createElement("dd");
            description.textContent = value;
            list.appendChild(description);
        });

        this.target.appendChild(list);
    }

    private createTable(headers: string[]): HTMLTableElement {
        const table = document.createElement("table");
        const thead = document.createElement("thead");
        const headerRow = document.createElement("tr");

        headers.forEach((label: string) => {
            const th = document.createElement("th");
            th.textContent = label;
            headerRow.appendChild(th);
        });

        thead.appendChild(headerRow);
        table.appendChild(thead);

        return table;
    }

    private appendTableRow(tbody: HTMLTableSectionElement, values: string[], className?: string): void {
        const row = document.createElement("tr");
        if (className) {
            row.className = className;
        }

        values.forEach((value: string) => {
            const cell = document.createElement("td");
            cell.textContent = value;
            row.appendChild(cell);
        });

        tbody.appendChild(row);
    }

    private appendMessage(message: string): void {
        const paragraph = document.createElement("p");
        paragraph.textContent = message;
        this.target.appendChild(paragraph);
    }

    private buildDatasetSignature(columns: DataViewMetadataColumn[], hasMoreData: boolean): string {
        const columnNames = columns.map((column: DataViewMetadataColumn) => column.queryName || column.displayName || "").join(",");
        const firstKey = this.accumulatedRows.length ? this.recordKey(this.accumulatedRows[0]) : "";
        const lastKey = this.accumulatedRows.length ? this.recordKey(this.accumulatedRows[this.accumulatedRows.length - 1]) : "";

        return [
            `rows=${this.accumulatedRows.length}`,
            `cols=${columns.length}`,
            `segment=${hasMoreData}`,
            columnNames,
            `first=${firstKey}`,
            `last=${lastKey}`
        ].join("|");
    }

    private buildIncomingDataSignature(records: SourceRecord[], columns: DataViewMetadataColumn[], hasMoreData: boolean): string {
        const columnNames = columns.map((column: DataViewMetadataColumn) => column.queryName || column.displayName || "").join(",");
        const firstKey = records.length ? this.recordKey(records[0]) : "";
        const lastKey = records.length ? this.recordKey(records[records.length - 1]) : "";

        return [
            `rows=${records.length}`,
            `cols=${columns.length}`,
            `segment=${hasMoreData}`,
            columnNames,
            `first=${firstKey}`,
            `last=${lastKey}`
        ].join("|");
    }

    private resetAccumulationIfDatasetChanged(
        records: SourceRecord[],
        columns: DataViewMetadataColumn[],
        hasMoreData: boolean
    ): void {
        const incomingSignature = this.buildIncomingDataSignature(records, columns, hasMoreData);
        const startsNewSegmentedLoad = this.allDataLoaded && hasMoreData && !this.isFetching;
        const incomingChangedAfterComplete = this.allDataLoaded &&
            !this.isFetching &&
            !!this.lastIncomingDataSignature &&
            incomingSignature !== this.lastIncomingDataSignature &&
            !this.recordsAlreadyAccumulated(records);

        if (startsNewSegmentedLoad || incomingChangedAfterComplete) {
            this.accumulatedRows = [];
            this.accumulatedRowKeys = new Set<string>();
            this.accumulatedRowCount = 0;
            this.fetchCount = 0;
            this.analyticsCache = null;
            this.detailCacheByUnit = new Map<string, UnitDetailRow[]>();
            this.lastDatasetSignature = "";
            this.segmentLoadStartTime = undefined;
            this.allDataLoaded = false;
        }

        this.lastIncomingDataSignature = incomingSignature;
    }

    private recordsAlreadyAccumulated(records: SourceRecord[]): boolean {
        return records.length > 0 && records.every((record: SourceRecord) => this.accumulatedRowKeys.has(this.recordKey(record)));
    }

    private hasMissingRole(columnIndexes: ColumnIndexes): boolean {
        return [
            columnIndexes.unidadGerencial,
            columnIndexes.region,
            columnIndexes.provincia,
            columnIndexes.distrito,
            columnIndexes.codigoLocal,
            columnIndexes.nombreLocal,
            columnIndexes.idSolicitud,
            columnIndexes.estadoSolicitud,
            columnIndexes.nivelRiesgo,
            columnIndexes.montoIntervencion,
            columnIndexes.beneficiarios,
            columnIndexes.montoAsignado,
            columnIndexes.montoTransferencia,
            columnIndexes.montoRetirado,
            columnIndexes.grupo,
            columnIndexes.cantidad
        ].some((index: number) => index < 0);
    }

    private describeDataReduction(table?: DataViewTable): string {
        if (!table?.columns.length) {
            return "(sin columnas)";
        }

        return `columns=${table.columns.length}, rows=${table.rows?.length || 0}, window=30000`;
    }

    private displayValue(value: PrimitiveValue): string {
        if (value === null) {
            return "(blank)";
        }

        if (value === undefined) {
            return "(undefined)";
        }

        if (value instanceof Date) {
            return value.toISOString();
        }

        return String(value);
    }

    private normalizeText(value: PrimitiveValue): string {
        return this.displayValue(value)
            .normalize("NFD")
            .replace(/[\u0300-\u036f]/g, "")
            .trim()
            .replace(/\s+/g, " ")
            .toUpperCase();
    }

    private readString(row: DataViewTableRow, index: number, fallback = ""): string {
        if (index < 0) {
            return fallback;
        }

        return this.displayValue(row[index]).trim();
    }

    private valueKey(value: PrimitiveValue): string {
        if (value instanceof Date) {
            return `date:${value.toISOString()}`;
        }

        return `${typeof value}:${String(value)}`;
    }

    private toNumber(value: PrimitiveValue): number {
        if (typeof value === "number") {
            return Number.isFinite(value) ? value : 0;
        }

        const parsed = Number(value);
        return Number.isFinite(parsed) ? parsed : 0;
    }

    private formatNumber(value: number): string {
        return this.formatCurrencyNoDecimals(value);
    }

    private formatCurrencyNoDecimals(value: number): string {
        return `S/ ${value.toLocaleString(undefined, {
            maximumFractionDigits: 0,
            minimumFractionDigits: 0
        })}`;
    }

    private formatDecimal(value: number): string {
        return value.toLocaleString(undefined, {
            maximumFractionDigits: 2,
            minimumFractionDigits: 2
        });
    }

    private formatInteger(value: number): string {
        return value.toLocaleString(undefined, {
            maximumFractionDigits: 0
        });
    }

    private formatPercent(value: number): string {
        return `${value.toFixed(1)}%`;
    }

    private formatMs(value: number): string {
        if (!Number.isFinite(value)) {
            return "n/a";
        }

        return `${value.toFixed(2)} ms`;
    }

    private calculateProcessingRate(rowCount: number, updateTotalMs: number): number {
        if (!rowCount || !updateTotalMs) {
            return 0;
        }

        return (rowCount / updateTotalMs) * 1000;
    }

    private calculateMsPer10k(rowCount: number, updateTotalMs: number): number {
        if (!rowCount) {
            return 0;
        }

        return (updateTotalMs / rowCount) * 10000;
    }

    private recordKey(record: SourceRecord): string {
        return [
            this.valueKey(record.region),
            this.valueKey(record.provincia),
            this.valueKey(record.distrito),
            this.valueKey(record.codigoLocal),
            this.valueKey(record.nombreLocal),
            this.valueKey(record.idSolicitud),
            this.valueKey(record.unidadGerencial),
            this.valueKey(record.estadoSolicitud),
            this.valueKey(record.nivelRiesgo),
            this.valueKey(record.grupo),
            record.montoIntervencion.toString(),
            record.beneficiarios.toString(),
            record.cantidad.toString(),
            record.montoAsignado.toString(),
            record.montoTransferencia.toString(),
            record.montoRetirado.toString(),
            record.montoInversion.toString(),
            record.latitud.toString(),
            record.longitud.toString()
        ].join("|");
    }
}
