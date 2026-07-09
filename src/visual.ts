/// <reference path="./typings.d.ts" />

"use strict";

import powerbi from "powerbi-visuals-api";
import { FormattingSettingsService } from "powerbi-visuals-utils-formattingmodel";
import * as L from "leaflet";
import type { Feature, FeatureCollection, Geometry } from "geojson";
import peruRegionGeoJson from "./../Departamento_Region.light.json";
import "leaflet/dist/leaflet.css";
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
type RiskLabel = RiskLevel | "Sin Clasificación";
type RiskMapView = "Total" | RiskLevel;

const riskMapViews: RiskMapView[] = ["Total", "Bajo", "Medio", "Alto", "Muy Alto"];
const neutralTotalRegionFill = "#94a3b8";
const neutralTotalRegionFillOpacity = 0.34;

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

interface RegionAggregate {
    region: string;
    totalColegios: number;
    totalProvincias: number;
    montoInversion: number;
    estudiantesBeneficiarios: number;
    numeroSolicitudes: number;
    riskBreakdown: RiskBuckets;
}

interface MapHeatScale {
    maxTotalColegios: number;
    riskHeatScores: Map<string, number>;
}

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

interface SvgIconShape {
    tag: "circle" | "path" | "rect";
    attrs: Record<string, string>;
}

type InternalFilterKey = "region" | "provincia" | "distrito" | "codigoLocal" | "nombreLocal";

interface InternalFilters {
    region: string;
    provincia: string;
    distrito: string;
    codigoLocal: string;
    nombreLocal: string;
}

interface InternalFilterDefinition {
    key: InternalFilterKey;
    label: string;
    placeholder: string;
}

interface UnitDetailRow {
    unit: UnitKpiName;
    codigoLocal: string;
    nombreLocal: string;
    nivelRiesgo: string;
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

interface SchoolSummaryRow {
    codigoLocal: string;
    nombreLocal: string;
    nivelRiesgo: string;
    region: string;
    provincia: string;
    distrito: string;
    units: Set<UnitKpiName>;
}

interface RegionTooltipSummary {
    colegios: number;
    provincias: number;
}

type MapDrillMode = "provincias" | "distritos" | "colegios";

interface MapPointSummary {
    key: string;
    label: string;
    count: number;
    lat: number;
    lng: number;
    riskBreakdown: RiskBuckets;
}

interface SchoolCluster {
    key: string;
    count: number;
    lat: number;
    lng: number;
    labels: string[];
    riskBreakdown: RiskBuckets;
}

interface MapPointAccumulator {
    key: string;
    label: string;
    regionKey: string;
    count: number;
    latTotal: number;
    lngTotal: number;
    riskBreakdown: RiskBuckets;
}

interface SchoolPointIndex {
    regionKey: string;
    label: string;
    lat: number;
    lng: number;
    riskLevel: string;
}

interface StreamingVisualIndexes {
    regionAnalytics: Map<string, AnalyticsEngine>;
    provinciaPoints: Map<string, MapPointAccumulator>;
    distritoPoints: Map<string, MapPointAccumulator>;
    schoolPoints: SchoolPointIndex[];
    detailRowsByUnit: Map<UnitKpiName, Map<string, UnitDetailRow>>;
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
    riesgoLabel: RiskLabel;
    colegioLabel: string;
}

interface PerformanceMetrics {
    updateTotalMs: number;
    firstUpdateWaitMs?: number;
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

interface AnalyticsCacheEntry {
    analytics: AnalyticsEngine;
    datasetSignature: string;
    rowCount: number;
    lastUsed: number;
}

function measure<T>(label: string, fn: () => T): { value: T; ms: number } {
    const start = performance.now();
    const value = fn();
    const end = performance.now();
    return { value, ms: end - start };
}

function getRiskGradientColor(score: number): string {
    if (!Number.isFinite(score) || score < 0) {
        return "#e5e7eb";
    }

    const stops = [
        { at: 0, color: [253, 224, 71] },
        { at: 0.32, color: [250, 204, 21] },
        { at: 0.55, color: [251, 146, 60] },
        { at: 0.78, color: [239, 68, 68] },
        { at: 1, color: [220, 38, 38] }
    ];
    const clamped = Math.max(0, Math.min(1, score));

    for (let index = 1; index < stops.length; index += 1) {
        const previous = stops[index - 1];
        const next = stops[index];

        if (clamped <= next.at) {
            const localRatio = (clamped - previous.at) / (next.at - previous.at);
            const color = previous.color.map((channel: number, channelIndex: number) => (
                Math.round(channel + (next.color[channelIndex] - channel) * localRatio)
            ));
            return `rgb(${color[0]}, ${color[1]}, ${color[2]})`;
        }
    }

    return "rgb(239, 68, 68)";
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

    public static build(records: Iterable<SourceRecord>): AnalyticsEngine {
        const engine = AnalyticsEngine.createEmpty();
        engine.addRecords(records);
        engine.finalizeBuild();

        return engine;
    }

    public static createEmpty(): AnalyticsEngine {
        return new AnalyticsEngine(AnalyticsEngine.createResult());
    }

    public addRecords(records: Iterable<SourceRecord>): void {
        let flatIndexesMs = 0;
        let relationIndexesMs = 0;

        for (const record of records) {
            const recordKeys = AnalyticsEngine.getRecordKeys(record);

            const flatStart = performance.now();
            AnalyticsEngine.updateBucket(this.result.totals, recordKeys, record);
            AnalyticsEngine.updateMapBucket(this.result.byRegion, recordKeys.regionKey, recordKeys.regionLabel, recordKeys, record);
            AnalyticsEngine.updateMapBucket(this.result.byProvincia, recordKeys.provinciaKey, recordKeys.provinciaLabel, recordKeys, record);
            AnalyticsEngine.updateMapBucket(this.result.byDistrito, recordKeys.distritoKey, recordKeys.distritoLabel, recordKeys, record);
            AnalyticsEngine.updateMapBucket(this.result.byUnidad, recordKeys.unidadKey, recordKeys.unidadLabel, recordKeys, record);
            AnalyticsEngine.updateMapBucket(this.result.byEstado, recordKeys.estadoKey, recordKeys.estadoLabel, recordKeys, record);
            AnalyticsEngine.updateMapBucket(this.result.byRiesgo, recordKeys.riesgoKey, recordKeys.riesgoLabel, recordKeys, record);
            AnalyticsEngine.updateMapBucket(this.result.byColegio, recordKeys.colegioKey, recordKeys.colegioLabel, recordKeys, record);
            AnalyticsEngine.updateUnitKpiBucket(this.result.byUnidadKpi, record);
            flatIndexesMs += performance.now() - flatStart;

            const relationStart = performance.now();
            AnalyticsEngine.addRelation(this.result.provinciasByRegion, recordKeys.regionKey, recordKeys.provinciaKey);
            AnalyticsEngine.addRelation(this.result.distritosByProvincia, recordKeys.provinciaKey, recordKeys.distritoKey);
            AnalyticsEngine.addRelation(this.result.colegiosByDistrito, recordKeys.distritoKey, recordKeys.colegioKey);
            relationIndexesMs += performance.now() - relationStart;
        }

        this.flatIndexesMs += flatIndexesMs;
        this.relationIndexesMs += relationIndexesMs;
    }

    public finalizeBuild(): void {
        this.buildFlatSummaries();
        this.lazySampleQueryMs = 0;
    }

    public buildFlatSummaries(): void {
        this.regionSummaries = this.mapToSummaries(this.result.byRegion, "region");
        this.unitSummaries = this.mapToSummaries(this.result.byUnidad);
        this.stateSummaries = [];
        this.riskSummaries = [];
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
        if (!this.stateSummaries.length && this.result.byEstado.size) {
            this.stateSummaries = this.mapToSummaries(this.result.byEstado);
        }

        return this.stateSummaries;
    }

    public getRiesgos(): RiskLevelSummary[] {
        if (!this.riskSummaries.length && this.result.byRiesgo.size) {
            this.riskSummaries = this.buildRiskSummaries();
        }

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
        } else if (recordKeys.riesgoLabel === "Muy Alto") {
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
            porcentajeCritico: this.getCriticalPercent(bucket)
        };
    }

    private static normalizeRiskLevel(value: PrimitiveValue): RiskLabel {
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

        if (normalized === "bajo") {
            return "Bajo";
        }

        if (normalized === "sin clasificación" || normalized === "sinclasificacion" || normalized === "sin clasificacion") {
            return "Sin Clasificación";
        }

        return "Sin Clasificación";
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
    private streamingIndexes: StreamingVisualIndexes = this.createStreamingIndexes();
    private lastFetchClickTime?: number;
    private isFetching = false;
    private allDataLoaded = false;
    private accumulatedRowCount = 0;
    private fetchCount = 0;
    private analyticsCache: AnalyticsEngine | null = null;
    private incrementalAnalytics: AnalyticsEngine | null = null;
    private analyticsCacheBySignature: Map<string, AnalyticsCacheEntry> = new Map<string, AnalyticsCacheEntry>();
    private selectedRegionAnalyticsCache: AnalyticsEngine | null = null;
    private selectedRegionAnalyticsCacheKey = "";
    private lastDatasetSignature = "";
    private lastIncomingDataSignature = "";
    private lastAccumulatedIncomingSignature = "";
    private lastColumnSignature = "";
    private segmentLoadStartTime?: number;
    private readonly visualCreatedAt = performance.now();
    private firstUpdateCaptured = false;
    private updateCount = 0;
    private cacheUseSequence = 0;
    private currentView: VisualView = "summary";
    private selectedRiskView: RiskMapView = "Total";
    private mapRankingLimit = 26;
    private selectedRegion: string | null = null;
    private internalFilters: InternalFilters = {
        region: "",
        provincia: "",
        distrito: "",
        codigoLocal: "",
        nombreLocal: ""
    };
    private dimensionStringPool: Map<string, string> = new Map<string, string>();
    private openInternalFilter: InternalFilterKey | null = null;
    private internalFilterAnalyticsCache: Map<string, AnalyticsEngine> = new Map<string, AnalyticsEngine>();
    private autoFocusedRegion: string | null = null;
    private suppressedAutoFocusSignature = "";
    private leafletMap: L.Map | null = null;
    private activeMapDrillLayer: L.LayerGroup | null = null;
    private activeMapDrillKey = "";
    private mapContextMenu: HTMLElement | null = null;
    private latestDiagnostics: TableDiagnostics | null = null;
    private detailModalUnit: UnitKpiName | null = null;
    private isDetailModalOpen = false;
    private isSchoolListModalOpen = false;
    private schoolListRegionFilter: string | null = null;
    private selectedSchoolKey: string | null = null;
    private selectedSchoolTab: UnitKpiName = "UGRD";
    private detailSearchText = "";
    private detailPage = 1;
    private detailPageSize = 20;
    private detailCacheByUnit: Map<string, UnitDetailRow[]> = new Map<string, UnitDetailRow[]>();
    private readonly maxAnalyticsCacheEntries = 4;
    private readonly maxDetailCacheEntries = 1;
    private readonly maxSchoolPointIndexes = 30000;
    private readonly enableAnalyticsCache = false;
    private readonly enableInternalFilterAnalyticsCache = false;
    private readonly enableInterimLoadingRender = true;
    private readonly enableSelectedRegionCache = false;
    private readonly enableSegmentAutoFetch = true;
    private readonly useCurrentDataViewOnly = false;
    private formattingSettings: VisualFormattingSettingsModel = new VisualFormattingSettingsModel();
    private formattingSettingsService: FormattingSettingsService;
    private handleKeydown = (event: KeyboardEvent): void => {
        if (event.key !== "Escape") {
            return;
        }

        if (this.selectedSchoolKey) {
            this.closeSchoolInterventionsModal();
            return;
        }

        if (this.isSchoolListModalOpen) {
            this.closeSchoolListModal();
            return;
        }

        if (this.isDetailModalOpen) {
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
        this.removeLeafletMap();
        this.disposeHeavyState();
        this.target.replaceChildren();
        window.removeEventListener("keydown", this.handleKeydown);
    }

    public update(options: VisualUpdateOptions): void {
        this.updateCount += 1;
        const updateStart = performance.now();
        const firstUpdateWaitMs = this.firstUpdateCaptured
            ? undefined
            : updateStart - this.visualCreatedAt;
        this.firstUpdateCaptured = true;
        const fetchWaitMs = this.lastFetchClickTime
            ? performance.now() - this.lastFetchClickTime
            : undefined;
        this.lastFetchClickTime = undefined;

        if (this.canReuseDiagnosticsForResize(options)) {
            const diagnostics = this.latestDiagnostics as TableDiagnostics;
            const renderStart = performance.now();
            diagnostics.performanceMetrics.analyticsCacheHit = true;
            diagnostics.performanceMetrics.analyticsRebuilt = false;
            diagnostics.performanceMetrics.buildRecordsMs = 0;
            diagnostics.performanceMetrics.accumulateRowsMs = 0;
            diagnostics.performanceMetrics.buildAnalyticsEngineMs = 0;
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
            return;
        }

        const dataViews: DataView[] = options.dataViews || [];
        const dataView = dataViews[0];

        if (dataView) {
            this.formattingSettings = this.formattingSettingsService.populateFormattingSettingsModel(
                VisualFormattingSettingsModel,
                dataView
            );
        }

        const diagnostics = this.buildDiagnostics(dataViews, fetchWaitMs, options.operationKind, firstUpdateWaitMs);
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
    }

    private renderCurrentView(): void {
        if (this.latestDiagnostics) {
            this.render(this.latestDiagnostics);
        }
    }

    private canReuseDiagnosticsForResize(options: VisualUpdateOptions): boolean {
        const updateType = options.type || 0;
        const resizeType = powerbi.VisualUpdateType.Resize;
        const isResizeOnly = updateType === resizeType;

        return isResizeOnly &&
            !!this.latestDiagnostics &&
            !this.latestDiagnostics.isLoading &&
            !!this.latestDiagnostics.analytics;
    }

    public getFormattingModel(): powerbi.visuals.FormattingModel {
        return this.formattingSettingsService.buildFormattingModel(this.formattingSettings);
    }

    private buildDiagnostics(
        dataViews: DataView[],
        fetchWaitMs?: number,
        operationKind?: powerbi.VisualDataChangeOperationKind,
        firstUpdateWaitMs?: number
    ): TableDiagnostics {
        const performanceMetrics: PerformanceMetrics = {
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
            fetchWaitMs: fetchWaitMs,
            updateTotalMs: 0,
            segmentLoadMs: 0,
            readDataViewMs: 0,
            findColumnsMs: 0
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
        this.resetAccumulationIfDatasetChanged(records, columns, hasMoreData, operationKind);

        if (this.isFetching && fetchWaitMs !== undefined) {
            this.isFetching = false;
        }

        this.requestMoreDataIfNeeded(hasMoreData, performanceMetrics);

        const incomingDataSignature = this.lastIncomingDataSignature;
        const accumulateResult = measure("accumulateRows", () => this.accumulateRecords(records, hasMoreData, incomingDataSignature));
        performanceMetrics.accumulateRowsMs = accumulateResult.ms;
        this.accumulatedRowCount = this.accumulatedRowCount || records.length;
        if (hasMoreData && this.enableSegmentAutoFetch) {
            const interimAnalyticsResult = this.enableInterimLoadingRender && this.accumulatedRowCount
                ? measure("buildInterimAnalyticsEngine", () => this.incrementalAnalytics || AnalyticsEngine.build(this.accumulatedRows))
                : null;
            const interimAnalytics = interimAnalyticsResult?.value || null;
                if (interimAnalyticsResult) {
                performanceMetrics.buildAnalyticsEngineMs = interimAnalyticsResult.ms;
                const interimBuildMetrics = interimAnalyticsResult.value.getBuildMetrics();
                performanceMetrics.buildFlatIndexesMs = interimBuildMetrics.buildFlatIndexesMs;
                performanceMetrics.buildRelationIndexesMs = interimBuildMetrics.buildRelationIndexesMs;
                performanceMetrics.lazySampleQueryMs = interimBuildMetrics.lazySampleQueryMs;
            }

            return {
                dataViewCount: dataViews.length,
                hasTable: !!table,
                columnIndexes,
                columns,
                segmentFilas: records.length,
                accumulatedFilas: this.accumulatedRowCount,
                hasMoreData,
                isLoading: true,
                fetchCount: this.fetchCount,
                datasetSignature: "",
                analyticsCacheHit: false,
                analyticsRebuilt: !!interimAnalytics,
                dataReductionText,
                analytics: interimAnalytics,
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

        let analytics: AnalyticsEngine | null = null;
        let analyticsCacheHit = false;
        let analyticsRebuilt = false;
        const cachedAnalytics = this.enableAnalyticsCache
            ? this.getCachedAnalytics(datasetSignature)
            : null;

        if (cachedAnalytics) {
            analytics = cachedAnalytics;
            analyticsCacheHit = true;
        } else if (this.incrementalAnalytics) {
            const analyticsResult = measure("finalizeIncrementalAnalyticsEngine", () => {
                this.incrementalAnalytics?.finalizeBuild();
                return this.incrementalAnalytics as AnalyticsEngine;
            });
            performanceMetrics.buildAnalyticsEngineMs = analyticsResult.ms;
            analytics = analyticsResult.value;
            if (this.enableAnalyticsCache) {
                this.cacheAnalytics(datasetSignature, analytics, this.accumulatedRowCount);
            }
            analyticsRebuilt = true;

            const buildMetrics = analytics.getBuildMetrics();
            performanceMetrics.buildFlatIndexesMs = buildMetrics.buildFlatIndexesMs;
            performanceMetrics.buildRelationIndexesMs = buildMetrics.buildRelationIndexesMs;
            performanceMetrics.lazySampleQueryMs = buildMetrics.lazySampleQueryMs;
        } else {
            const analyticsResult = measure("buildAnalyticsEngine", () => AnalyticsEngine.build(this.accumulatedRows));
            performanceMetrics.buildAnalyticsEngineMs = analyticsResult.ms;
            analytics = analyticsResult.value;
            if (this.enableAnalyticsCache) {
                this.cacheAnalytics(datasetSignature, analytics, this.accumulatedRowCount);
            }
            analyticsRebuilt = true;

            const buildMetrics = analytics.getBuildMetrics();
            performanceMetrics.buildFlatIndexesMs = buildMetrics.buildFlatIndexesMs;
            performanceMetrics.buildRelationIndexesMs = buildMetrics.buildRelationIndexesMs;
            performanceMetrics.lazySampleQueryMs = buildMetrics.lazySampleQueryMs;
        }

        this.analyticsCache = analytics;
        this.accumulatedRowKeys.clear();
        this.lastDatasetSignature = datasetSignature;
        const lazyDiagnostics = analytics.getLazyNavigationDiagnostics();
        performanceMetrics.indexedRegiones = lazyDiagnostics.regionesIndexadas;
        performanceMetrics.indexedProvincias = lazyDiagnostics.provinciasIndexadas;
        performanceMetrics.indexedDistritos = lazyDiagnostics.distritosIndexados;
        performanceMetrics.indexedColegios = lazyDiagnostics.colegiosIndexados;
        performanceMetrics.fetchCount = this.fetchCount;
        performanceMetrics.analyticsCacheHit = analyticsCacheHit;
        performanceMetrics.analyticsRebuilt = analyticsRebuilt;
        performanceMetrics.rowsPerSecond = this.calculateProcessingRate(this.accumulatedRowCount, performanceMetrics.buildAnalyticsEngineMs || performanceMetrics.updateTotalMs);
        performanceMetrics.msPer10kRows = this.calculateMsPer10k(this.accumulatedRowCount, performanceMetrics.buildAnalyticsEngineMs || performanceMetrics.updateTotalMs);

        return {
            dataViewCount: dataViews.length,
            hasTable: !!table,
            columnIndexes,
            columns,
            segmentFilas: records.length,
            accumulatedFilas: this.accumulatedRowCount,
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

    private requestMoreDataIfNeeded(hasMoreData: boolean, performanceMetrics: PerformanceMetrics): void {
        if (!hasMoreData || !this.enableSegmentAutoFetch) {
            return;
        }

        this.allDataLoaded = false;
        if (this.segmentLoadStartTime === undefined) {
            this.segmentLoadStartTime = performance.now();
        }

        if (!this.isFetching) {
            this.isFetching = true;
            this.fetchCount += 1;
            this.lastFetchClickTime = performance.now();
            this.host.fetchMoreData();
        }

        performanceMetrics.fetchCount = this.fetchCount;
        performanceMetrics.segmentLoadMs = performance.now() - this.segmentLoadStartTime;
    }

    private findColumnIndexByRole(columns: DataViewMetadataColumn[], roleName: string): number {
        return columns.findIndex((column: DataViewMetadataColumn) => !!column.roles?.[roleName]);
    }

    private buildRecords(rows: DataViewTableRow[], columnIndexes: ColumnIndexes): SourceRecord[] {
        if (this.hasMissingRole(columnIndexes)) {
            return [];
        }

        return rows.map((row: DataViewTableRow) => ({
            unidadGerencial: this.internDimensionString(this.readString(row, columnIndexes.unidadGerencial)),
            region: this.internDimensionString(this.readString(row, columnIndexes.region)),
            provincia: this.internDimensionString(this.readString(row, columnIndexes.provincia)),
            codigoLocal: this.readString(row, columnIndexes.codigoLocal),
            nombreLocal: this.readString(row, columnIndexes.nombreLocal, "-"),
            montoIntervencion: this.toNumber(row[columnIndexes.montoIntervencion]),
            beneficiarios: this.toNumber(row[columnIndexes.beneficiarios]),
            montoAsignado: this.toNumber(row[columnIndexes.montoAsignado]),
            montoTransferencia: this.toNumber(row[columnIndexes.montoTransferencia]),
            montoRetirado: this.toNumber(row[columnIndexes.montoRetirado]),
            grupo: this.internDimensionString(this.readString(row, columnIndexes.grupo)),
            cantidad: this.toNumber(row[columnIndexes.cantidad]),
            distrito: this.internDimensionString(this.readString(row, columnIndexes.distrito)),
            idSolicitud: this.readString(row, columnIndexes.idSolicitud),
            estadoSolicitud: this.internDimensionString(this.readString(row, columnIndexes.estadoSolicitud)),
            nivelRiesgo: this.internDimensionString(this.readString(row, columnIndexes.nivelRiesgo)),
            montoInversion: this.toNumber(row[columnIndexes.montoInversion]) || this.toNumber(row[columnIndexes.montoIntervencion]),
            latitud: this.toNumber(row[columnIndexes.latitud]),
            longitud: this.toNumber(row[columnIndexes.longitud])
        }));
    }

    private createStreamingIndexes(): StreamingVisualIndexes {
        return {
            regionAnalytics: new Map<string, AnalyticsEngine>(),
            provinciaPoints: new Map<string, MapPointAccumulator>(),
            distritoPoints: new Map<string, MapPointAccumulator>(),
            schoolPoints: [],
            detailRowsByUnit: new Map<UnitKpiName, Map<string, UnitDetailRow>>()
        };
    }

    private resetStreamingState(): void {
        this.accumulatedRows = [];
        this.accumulatedRowKeys = new Set<string>();
        this.streamingIndexes = this.createStreamingIndexes();
        this.accumulatedRowCount = 0;
        this.detailCacheByUnit = new Map<string, UnitDetailRow[]>();
        this.internalFilterAnalyticsCache = new Map<string, AnalyticsEngine>();
        this.dimensionStringPool = new Map<string, string>();
    }

    private disposeHeavyState(): void {
        this.resetStreamingState();
        this.analyticsCache = null;
        this.incrementalAnalytics = null;
        this.analyticsCacheBySignature = new Map<string, AnalyticsCacheEntry>();
        this.selectedRegionAnalyticsCache = null;
        this.selectedRegionAnalyticsCacheKey = "";
        this.latestDiagnostics = null;
        this.lastDatasetSignature = "";
        this.lastIncomingDataSignature = "";
        this.lastAccumulatedIncomingSignature = "";
        this.lastColumnSignature = "";
        this.segmentLoadStartTime = undefined;
        this.lastFetchClickTime = undefined;
        this.isFetching = false;
        this.allDataLoaded = false;
        this.fetchCount = 0;
        this.cacheUseSequence = 0;
        this.currentView = "summary";
        this.selectedRiskView = "Total";
        this.internalFilters = {
            region: "",
            provincia: "",
            distrito: "",
            codigoLocal: "",
            nombreLocal: ""
        };
        this.selectedRegion = null;
        this.openInternalFilter = null;
        this.autoFocusedRegion = null;
        this.suppressedAutoFocusSignature = "";
        this.activeMapDrillKey = "";
        this.mapContextMenu = null;
        this.detailModalUnit = null;
        this.isDetailModalOpen = false;
        this.isSchoolListModalOpen = false;
        this.schoolListRegionFilter = null;
        this.selectedSchoolKey = null;
        this.selectedSchoolTab = "UGRD";
        this.detailSearchText = "";
        this.detailPage = 1;
    }

    private accumulateRecords(records: SourceRecord[], hasMoreData: boolean, incomingDataSignature: string): void {
        if (this.useCurrentDataViewOnly) {
            this.resetStreamingState();
            this.incrementalAnalytics = AnalyticsEngine.createEmpty();
            this.incrementalAnalytics.addRecords(records);
            this.updateStreamingIndexes(records);
            this.accumulatedRows = records.slice();
            this.accumulatedRowCount = records.length;
            this.lastAccumulatedIncomingSignature = incomingDataSignature;
            return;
        }

        if (incomingDataSignature && incomingDataSignature === this.lastAccumulatedIncomingSignature) {
            return;
        }

        if (!hasMoreData && !this.isFetching && this.fetchCount === 0 && !this.accumulatedRowCount) {
            this.resetStreamingState();
            this.incrementalAnalytics = AnalyticsEngine.createEmpty();
            this.incrementalAnalytics.addRecords(records);
            this.updateStreamingIndexes(records);
            this.accumulatedRows = records.slice();
            this.accumulatedRowCount = records.length;
            this.lastAccumulatedIncomingSignature = incomingDataSignature;
            return;
        }

        const appendedRecords: SourceRecord[] = [];
        records.forEach((record: SourceRecord) => {
            const recordKey = this.recordKey(record);
            if (this.accumulatedRowKeys.has(recordKey)) {
                return;
            }

            this.accumulatedRowKeys.add(recordKey);
            appendedRecords.push(record);
        });

        if (appendedRecords.length) {
            if (!this.incrementalAnalytics) {
                this.incrementalAnalytics = AnalyticsEngine.createEmpty();
            }

            this.incrementalAnalytics.addRecords(appendedRecords);
            this.updateStreamingIndexes(appendedRecords);
            this.accumulatedRows.push(...appendedRecords);
            this.accumulatedRowCount += appendedRecords.length;
        }

        this.lastAccumulatedIncomingSignature = incomingDataSignature;
    }

    private updateStreamingIndexes(records: SourceRecord[]): void {
        records.forEach((record: SourceRecord) => {
            this.updateRegionAnalyticsIndex(record);
            this.updateMapPointIndexes(record);
            this.updateDetailIndex(record);
        });

        this.detailCacheByUnit = new Map<string, UnitDetailRow[]>();
    }

    private updateRegionAnalyticsIndex(record: SourceRecord): void {
        const regionKey = this.normalizeRegionForMap(record.region);
        if (!regionKey) {
            return;
        }

        const keys = Array.from(new Set<string>([regionKey, ...this.getRegionAliases(regionKey)]));
        keys.forEach((key: string) => {
            const regionEngine = this.streamingIndexes.regionAnalytics.get(key) || AnalyticsEngine.createEmpty();
            regionEngine.addRecords([record]);
            this.streamingIndexes.regionAnalytics.set(key, regionEngine);
        });
    }

    private updateMapPointIndexes(record: SourceRecord): void {
        if (!this.hasValidCoordinates(record)) {
            return;
        }

        const regionKey = this.normalizeRegionForMap(record.region);
        this.updateMapPointAccumulator(this.streamingIndexes.provinciaPoints, regionKey, record.provincia, record);
        this.updateMapPointAccumulator(this.streamingIndexes.distritoPoints, regionKey, record.distrito, record);
        if (this.streamingIndexes.schoolPoints.length >= this.maxSchoolPointIndexes) {
            return;
        }

        this.streamingIndexes.schoolPoints.push({
            regionKey,
            label: record.codigoLocal || record.nombreLocal || "-",
            lat: record.latitud,
            lng: record.longitud,
            riskLevel: record.nivelRiesgo
        });
    }

    private updateMapPointAccumulator(
        target: Map<string, MapPointAccumulator>,
        regionKey: string,
        labelValue: string,
        record: SourceRecord
    ): void {
        const label = labelValue || "-";
        const key = `${regionKey}|${this.normalizeText(label)}`;
        const accumulator = target.get(key) || {
            key,
            label,
            regionKey,
            count: 0,
            latTotal: 0,
            lngTotal: 0,
            riskBreakdown: { bajo: 0, medio: 0, alto: 0, muyAlto: 0 }
        };

        accumulator.count += 1;
        accumulator.latTotal += record.latitud;
        accumulator.lngTotal += record.longitud;
        this.addRiskToBreakdown(accumulator.riskBreakdown, record.nivelRiesgo);
        target.set(key, accumulator);
    }

    private updateDetailIndex(record: SourceRecord): void {
        const unit = this.normalizeText(record.unidadGerencial) as UnitKpiName;
        if (!this.isUnitKpiName(unit)) {
            return;
        }

        const unitRows = this.streamingIndexes.detailRowsByUnit.get(unit) || new Map<string, UnitDetailRow>();
        const codigoKey = this.normalizeText(record.codigoLocal);
        const grupoKey = unit === "UGME" ? this.normalizeText(record.grupo) : "";
        const detailKey = unit === "UGME" ? `${codigoKey}|${grupoKey}` : codigoKey;
        const row = unitRows.get(detailKey) || {
            unit,
            codigoLocal: record.codigoLocal || "-",
            nombreLocal: record.nombreLocal || "-",
            nivelRiesgo: record.nivelRiesgo || "-",
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

        unitRows.set(detailKey, row);
        this.streamingIndexes.detailRowsByUnit.set(unit, unitRows);
    }

    private isUnitKpiName(value: string): value is UnitKpiName {
        return value === "UGRD" || value === "UGEO" || value === "UGSC" || value === "UGM" || value === "UGME";
    }

    private render(diagnostics: TableDiagnostics): void {
        this.removeLeafletMap();
        this.target.replaceChildren();

        if (!diagnostics.hasTable) {
            this.appendMessage("No table DataView received.");
            return;
        }

        if (this.hasMissingRole(diagnostics.columnIndexes)) {
            this.appendMessage("Missing one or more required roles for unit KPI cards.");
            return;
        }

        if (diagnostics.isLoading && diagnostics.analytics) {
            this.reconcileInternalFilters();
            const renderAnalytics = this.getInternalFilteredAnalytics(diagnostics.analytics);
            const mapAnalytics = this.getInternalFilteredAnalytics(diagnostics.analytics, false);
            if (this.currentView === "ugme") {
                this.appendUgmeView(renderAnalytics, mapAnalytics);
            } else {
                this.appendSummaryView(renderAnalytics, mapAnalytics);
            }

            return;
        }

        if (diagnostics.isLoading) {
            this.appendFinalLoadingPanel(diagnostics);
            return;
        }

        if (!diagnostics.analytics) {
            this.appendMessage("Analytics Engine is not available.");
            return;
        }

        this.reconcileInternalFilters();
        const renderAnalytics = this.getInternalFilteredAnalytics(diagnostics.analytics);
        const mapAnalytics = this.getInternalFilteredAnalytics(diagnostics.analytics, false);
        if (this.currentView === "ugme") {
            this.appendUgmeView(renderAnalytics, mapAnalytics);
        } else {
            this.appendSummaryView(renderAnalytics, mapAnalytics);
        }

        if (this.isDetailModalOpen && this.detailModalUnit) {
            this.appendDetailModal(this.detailModalUnit);
        }

        if (this.isSchoolListModalOpen) {
            this.appendSchoolListModal();
        }

        if (this.selectedSchoolKey) {
            this.appendSchoolInterventionsModal(this.selectedSchoolKey);
        }
    }

    private appendSummaryView(engine: AnalyticsEngine, mapEngine: AnalyticsEngine = engine): void {
        const view = document.createElement("section");
        view.className = "dashboard-root with-internal-filters";
        this.appendInternalFilterBar(view);

        const layout = document.createElement("div");
        layout.className = "main-dashboard-layout";

        const mapPanel = document.createElement("section");
        mapPanel.className = "map-panel";
        this.appendPeruRiskMap(mapPanel, mapEngine);
        layout.appendChild(mapPanel);

        const unitsPanel = this.createSummaryUnitsPanel(engine);

        layout.appendChild(unitsPanel);
        view.appendChild(layout);

        this.target.appendChild(view);
    }

    private createSummaryUnitsPanel(engine: AnalyticsEngine): HTMLElement {
        const unitsPanel = document.createElement("section");
        unitsPanel.className = "units-panel";
        const unitEngine = this.getSelectedRegionEngine(engine);

        this.appendInformationBanner(unitsPanel, unitEngine);

        const header = document.createElement("div");
        header.className = "units-panel-header";

        const title = document.createElement("h1");
        title.textContent = this.selectedRegion
            ? `RESUMEN POR UNIDADES GERENCIALES - ${this.selectedRegion}`
            : "RESUMEN POR UNIDADES GERENCIALES";
        header.appendChild(title);

        const detailButton = document.createElement("button");
        detailButton.className = "units-detail-button";
        detailButton.type = "button";
        detailButton.textContent = "Ver detalle";
        detailButton.onclick = () => this.openSchoolListModal(this.selectedRegion);
        header.appendChild(detailButton);
        unitsPanel.appendChild(header);

        const grid = document.createElement("div");
        grid.className = "unit-grid";

        (["UGRD", "UGSC", "UGEO", "UGM"] as UnitKpiName[]).forEach((unit: UnitKpiName) => {
            grid.appendChild(this.createUnitCard(unitEngine.getUnitKpi(unit), false));
        });

        unitsPanel.appendChild(grid);

        const nextButton = document.createElement("button");
        nextButton.className = "floating-next-button";
        nextButton.type = "button";
        nextButton.textContent = ">";
        nextButton.onclick = () => this.switchVisualView("ugme");
        unitsPanel.appendChild(nextButton);

        return unitsPanel;
    }

    private switchVisualView(nextView: VisualView): void {
        this.currentView = nextView;
        if (!this.replaceUnitsPanelForCurrentView()) {
            this.renderCurrentView();
        }
    }

    private replaceUnitsPanelForCurrentView(): boolean {
        if (!this.latestDiagnostics?.analytics) {
            return false;
        }

        const layout = this.target.querySelector(".main-dashboard-layout");
        const currentPanel = layout?.querySelector(".units-panel");
        const root = this.target.querySelector(".dashboard-root");
        if (!layout || !currentPanel || !root) {
            return false;
        }

        this.reconcileInternalFilters();
        const engine = this.getInternalFilteredAnalytics(this.latestDiagnostics.analytics);
        const nextPanel = this.currentView === "ugme"
            ? this.createUgmeUnitsPanel(engine)
            : this.createSummaryUnitsPanel(engine);

        currentPanel.replaceWith(nextPanel);
        root.classList.toggle("ugme-dashboard-view", this.currentView === "ugme");
        return true;
    }

    private appendInternalFilterBar(container: HTMLElement): void {
        const filterBar = document.createElement("section");
        filterBar.className = "internal-filter-bar";

        const definitions: InternalFilterDefinition[] = [
            { key: "region", label: "Región", placeholder: "Todas" },
            { key: "provincia", label: "Provincia", placeholder: "Todas" },
            { key: "distrito", label: "Distrito", placeholder: "Todos" },
            { key: "codigoLocal", label: "Código local", placeholder: "Todos" }
        ];

        definitions.push({ key: "nombreLocal", label: "Nombre local", placeholder: "Todos" });

        definitions.forEach((definition: InternalFilterDefinition) => {
            filterBar.appendChild(this.createInternalFilterDropdown(definition));
        });

        const clearButton = document.createElement("button");
        clearButton.className = "internal-filter-clear";
        clearButton.type = "button";
        clearButton.textContent = "Limpiar";
        clearButton.disabled = !this.hasInternalFilters();
        clearButton.onclick = () => {
            this.internalFilters = {
                region: "",
                provincia: "",
                distrito: "",
                codigoLocal: "",
                nombreLocal: ""
            };
            this.openInternalFilter = null;
            this.detailPage = 1;
            this.internalFilterAnalyticsCache = new Map<string, AnalyticsEngine>();
            this.clearInteractiveSelectionState();
            this.renderCurrentView();
        };
        filterBar.appendChild(clearButton);

        container.appendChild(filterBar);
    }

    private createInternalFilterDropdown(definition: InternalFilterDefinition): HTMLElement {
        const wrapper = document.createElement("div");
        wrapper.className = `internal-filter-dropdown${this.openInternalFilter === definition.key ? " is-open" : ""}`;

        const label = document.createElement("span");
        label.className = "internal-filter-label";
        label.textContent = definition.label;
        wrapper.appendChild(label);

        const button = document.createElement("button");
        button.className = "internal-filter-button";
        button.type = "button";
        button.textContent = this.internalFilters[definition.key] || definition.placeholder;
        button.title = this.internalFilters[definition.key] || definition.placeholder;
        button.onclick = () => {
            this.openInternalFilter = this.openInternalFilter === definition.key ? null : definition.key;
            const filterBar = button.closest(".internal-filter-bar");
            filterBar?.querySelectorAll(".internal-filter-dropdown").forEach((dropdown: Element) => {
                dropdown.classList.remove("is-open");
            });
            if (this.openInternalFilter === definition.key) {
                wrapper.classList.add("is-open");
                const searchInput = wrapper.querySelector(".internal-filter-search input") as HTMLInputElement | null;
                window.setTimeout(() => searchInput?.focus(), 0);
            }
        };
        wrapper.appendChild(button);

        {
            const menu = document.createElement("div");
            menu.className = "internal-filter-menu";

            const searchWrap = document.createElement("div");
            searchWrap.className = "internal-filter-search";

            const searchIcon = document.createElement("span");
            searchIcon.textContent = "🔍";
            searchWrap.appendChild(searchIcon);

            const searchInput = document.createElement("input");
            searchInput.type = "search";
            searchInput.placeholder = `Buscar ${definition.label.toLowerCase()}`;
            searchInput.setAttribute("aria-label", `Buscar ${definition.label.toLowerCase()}`);
            searchWrap.appendChild(searchInput);
            menu.appendChild(searchWrap);

            const options = this.getInternalFilterOptions(definition.key);
            const list = document.createElement("div");
            list.className = "internal-filter-option-list";
            menu.appendChild(list);

            const renderOptions = (query: string): void => {
                const normalizedQuery = this.normalizeText(query);
                const filteredOptions = normalizedQuery
                    ? options.filter((option: string) => this.normalizeText(option).includes(normalizedQuery))
                    : options;
                const visibleOptions = filteredOptions.slice(0, 300);
                list.replaceChildren();
                list.appendChild(this.createInternalFilterOption(definition.key, "", definition.placeholder, true));
                visibleOptions.forEach((option: string) => {
                    list.appendChild(this.createInternalFilterOption(definition.key, option, option, false));
                });

                if (filteredOptions.length > visibleOptions.length) {
                    const hint = document.createElement("div");
                    hint.className = "internal-filter-hint";
                    hint.textContent = `Mostrando ${visibleOptions.length} de ${filteredOptions.length}. Escribe para reducir.`;
                    list.appendChild(hint);
                }
            };

            searchInput.oninput = () => renderOptions(searchInput.value);
            searchInput.onclick = (event: MouseEvent) => event.stopPropagation();
            renderOptions("");

            wrapper.appendChild(menu);
        }

        return wrapper;
    }

    private createInternalFilterOption(
        key: InternalFilterKey,
        value: string,
        label: string,
        isAllOption: boolean
    ): HTMLButtonElement {
        const option = document.createElement("button");
        option.className = `internal-filter-option${this.internalFilters[key] === value ? " is-selected" : ""}`;
        option.type = "button";
        option.textContent = label;
        option.title = label;
        option.onclick = () => {
            this.setInternalFilter(key, value);
        };

        if (isAllOption) {
            option.classList.add("is-all-option");
        }

        return option;
    }

    private setInternalFilter(key: InternalFilterKey, value: string): void {
        this.internalFilters = {
            ...this.internalFilters,
            [key]: value
        };

        if (key === "region") {
            this.internalFilters.provincia = "";
            this.internalFilters.distrito = "";
            this.internalFilters.codigoLocal = "";
            this.internalFilters.nombreLocal = "";
        } else if (key === "provincia") {
            this.internalFilters.distrito = "";
            this.internalFilters.codigoLocal = "";
            this.internalFilters.nombreLocal = "";
        } else if (key === "distrito") {
            this.internalFilters.codigoLocal = "";
            this.internalFilters.nombreLocal = "";
        } else if (key === "codigoLocal") {
            this.internalFilters.nombreLocal = "";
        }

        this.openInternalFilter = null;
        this.detailPage = 1;
        this.internalFilterAnalyticsCache = new Map<string, AnalyticsEngine>();
        this.clearInteractiveSelectionState();
        this.renderCurrentView();
    }

    private clearInteractiveSelectionState(): void {
        this.selectedRegion = null;
        this.autoFocusedRegion = null;
        this.suppressedAutoFocusSignature = "";
        this.selectedRegionAnalyticsCache = null;
        this.selectedRegionAnalyticsCacheKey = "";
        this.activeMapDrillKey = "";
    }

    private hasInternalFilters(): boolean {
        return !!(
            this.internalFilters.region ||
            this.internalFilters.provincia ||
            this.internalFilters.distrito ||
            this.internalFilters.codigoLocal ||
            this.internalFilters.nombreLocal
        );
    }

    private hasRiskFilter(): boolean {
        return this.selectedRiskView !== "Total";
    }

    private hasDashboardFilters(includeRiskFilter = true): boolean {
        return this.hasInternalFilters() || (includeRiskFilter && this.hasRiskFilter());
    }

    private getInternalFilterOptions(key: InternalFilterKey): string[] {
        const values = new Set<string>();

        this.accumulatedRows.forEach((record: SourceRecord) => {
            if (!this.recordMatchesDashboardFilters(record, key)) {
                return;
            }

            const value = record[key];
            if (value) {
                values.add(value);
            }
        });

        return Array.from(values).sort((left: string, right: string) => left.localeCompare(right));
    }

    private recordMatchesDashboardFilters(record: SourceRecord, ignoreKey?: InternalFilterKey, includeRiskFilter = true): boolean {
        const filters = this.internalFilters;

        if (includeRiskFilter && !this.recordMatchesSelectedRisk(record)) {
            return false;
        }

        if (ignoreKey !== "region" && filters.region && this.normalizeText(record.region) !== this.normalizeText(filters.region)) {
            return false;
        }

        if (ignoreKey !== "provincia" && filters.provincia && this.normalizeText(record.provincia) !== this.normalizeText(filters.provincia)) {
            return false;
        }

        if (ignoreKey !== "distrito" && filters.distrito && this.normalizeText(record.distrito) !== this.normalizeText(filters.distrito)) {
            return false;
        }

        if (ignoreKey !== "codigoLocal" && filters.codigoLocal && this.normalizeText(record.codigoLocal) !== this.normalizeText(filters.codigoLocal)) {
            return false;
        }

        if (ignoreKey !== "nombreLocal" && filters.nombreLocal && this.normalizeText(record.nombreLocal) !== this.normalizeText(filters.nombreLocal)) {
            return false;
        }

        return true;
    }

    private reconcileInternalFilters(): void {
        (["region", "provincia", "distrito", "codigoLocal", "nombreLocal"] as InternalFilterKey[]).forEach((key: InternalFilterKey) => {
            if (!this.internalFilters[key]) {
                return;
            }

            const exists = this.accumulatedRows.some((record: SourceRecord) => (
                this.normalizeText(record[key]) === this.normalizeText(this.internalFilters[key]) &&
                this.recordMatchesDashboardFilters(record, key)
            ));

            if (!exists) {
                this.internalFilters[key] = "";
            }
        });
    }

    private getInternalFilteredAnalytics(baseEngine: AnalyticsEngine, includeRiskFilter = true): AnalyticsEngine {
        if (!this.hasDashboardFilters(includeRiskFilter) || !this.accumulatedRows.length) {
            return baseEngine;
        }

        const filterKey = this.getInternalFilterCacheKey(includeRiskFilter);
        if (this.enableInternalFilterAnalyticsCache) {
            const cached = this.internalFilterAnalyticsCache.get(filterKey);
            if (cached) {
                return cached;
            }
        }

        const filteredRows = this.accumulatedRows.filter((record: SourceRecord) => this.recordMatchesDashboardFilters(record, undefined, includeRiskFilter));
        const filteredEngine = AnalyticsEngine.build(filteredRows);

        if (this.enableInternalFilterAnalyticsCache) {
            this.internalFilterAnalyticsCache.set(filterKey, filteredEngine);

            if (this.internalFilterAnalyticsCache.size > 1) {
                const oldestKey = this.internalFilterAnalyticsCache.keys().next().value;
                if (oldestKey) {
                    this.internalFilterAnalyticsCache.delete(oldestKey);
                }
            }
        }

        return filteredEngine;
    }

    private getInternalFilterCacheKey(includeRiskFilter = true): string {
        return [
            this.lastDatasetSignature,
            this.normalizeText(this.internalFilters.region),
            this.normalizeText(this.internalFilters.provincia),
            this.normalizeText(this.internalFilters.distrito),
            this.normalizeText(this.internalFilters.codigoLocal),
            this.normalizeText(this.internalFilters.nombreLocal),
            includeRiskFilter ? this.selectedRiskView : "Total"
        ].join("|");
    }

    private appendPeruRiskMap(container: HTMLElement, engine: AnalyticsEngine): void {
        const title = document.createElement("h2");
        title.className = "map-title";
        title.textContent = "MAPA DE RIESGO POR REGION";
        container.appendChild(title);

        const map = document.createElement("div");
        map.className = "peru-risk-map";

        const leafletContainer = document.createElement("div");
        leafletContainer.className = "leaflet-risk-map";
        map.appendChild(leafletContainer);

        this.appendMapResetButton(map);
        this.appendRiskToggle(map);
        this.appendMapRankingControl(map, engine);
        container.appendChild(map);

        window.setTimeout(() => this.renderLeafletPeruMap(leafletContainer, engine), 0);
    }

    private renderLeafletPeruMap(mapElement: HTMLElement, engine: AnalyticsEngine): void {
        const aggregates = this.getRegionAggregates(engine);
        const aggregateMap = this.createRegionAggregateMap(aggregates);
        this.autoFocusedRegion = this.selectedRegion
            ? null
            : this.getAutoFocusedRegion(engine);
        const mapHeatScale = this.buildMapHeatScale(aggregates);
        const map = L.map(mapElement, {
            zoomControl: true,
            dragging: true,
            scrollWheelZoom: true,
            doubleClickZoom: false,
            touchZoom: true,
            boxZoom: true,
            keyboard: true
        });

        this.leafletMap = map;
        this.ensureOpenStreetMapTileLayer(map);

        let selectedLayer: L.Layer | null = null;
        const hideContextMenu = () => this.hideMapContextMenu();
        const geoJsonLayer = L.geoJSON(peruRegionGeoJson as FeatureCollection<Geometry>, {
            style: (feature?: Feature<Geometry>) => this.getRegionLayerStyle(feature, aggregateMap, mapHeatScale),
            onEachFeature: (feature: Feature<Geometry>, layer: L.Layer) => {
                const regionName = this.getGeoJsonRegionName(feature);
                const aggregate = this.findRegionAggregate(regionName, aggregateMap);
                const isSelected = this.isSelectedRegion(regionName);

                if (isSelected) {
                    selectedLayer = layer;
                }

                layer.bindTooltip(this.createRegionTooltip(regionName, aggregate), {
                    sticky: true,
                    direction: "auto",
                    opacity: 0.95,
                    className: "region-leaflet-tooltip"
                });

                layer.on({
                    mouseover: () => {
                        const pathLayer = layer as L.Path;
                        pathLayer.setStyle({
                            fillOpacity: 0.78,
                            weight: 3
                        });
                        pathLayer.bringToFront();
                    },
                    mouseout: () => {
                        geoJsonLayer.resetStyle(layer as L.Path);
                        (layer as L.Path).bringToFront();
                    },
                    click: () => {
                        hideContextMenu();
                        if (this.isRegionNameSelected(regionName)) {
                            this.resetMapFocus(true);
                        } else {
                            this.selectedRegion = regionName;
                            this.autoFocusedRegion = null;
                            this.suppressedAutoFocusSignature = "";
                            this.selectedRegionAnalyticsCache = null;
                            this.selectedRegionAnalyticsCacheKey = "";
                        }
                        this.renderCurrentView();
                    },
                    dblclick: (event: L.LeafletMouseEvent) => {
                        L.DomEvent.preventDefault(event.originalEvent);
                        L.DomEvent.stopPropagation(event.originalEvent);
                        this.resetMapFocus(true);
                        this.renderCurrentView();
                    },
                    contextmenu: (event: L.LeafletMouseEvent) => {
                        L.DomEvent.preventDefault(event.originalEvent);
                        L.DomEvent.stopPropagation(event.originalEvent);
                        this.showMapContextMenu(mapElement, map, regionName, event);
                    }
                });
            }
        }).addTo(map);

        map.on("click", hideContextMenu);
        map.on("movestart", hideContextMenu);
        geoJsonLayer.bringToFront();
        const focusLayer = selectedLayer as L.FeatureGroup | L.Polygon | L.Polyline | null;
        const focusBounds = focusLayer && "getBounds" in focusLayer
            ? focusLayer.getBounds()
            : geoJsonLayer.getBounds();
        map.fitBounds(focusBounds, this.getMapFitBoundsOptions());
        map.invalidateSize();
    }

    private getMapFitBoundsOptions(): L.FitBoundsOptions {
        if (this.selectedRiskView !== "Total") {
            return {
                paddingTopLeft: [18, 24],
                paddingBottomRight: [330, 28]
            };
        }

        return { padding: [20, 20] };
    }

    private getAutoFocusedRegion(engine: AnalyticsEngine): string | null {
        if (this.suppressedAutoFocusSignature === this.lastDatasetSignature) {
            return null;
        }

        const regions = engine.getRegions();
        return regions.length === 1 ? regions[0].label : null;
    }

    private resetMapFocus(suppressAutoFocus: boolean): void {
        this.selectedRegion = null;
        this.autoFocusedRegion = null;
        this.activeMapDrillKey = "";
        this.selectedRegionAnalyticsCache = null;
        this.selectedRegionAnalyticsCacheKey = "";
        this.hideMapContextMenu();

        if (suppressAutoFocus) {
            this.suppressedAutoFocusSignature = this.lastDatasetSignature;
        }
    }

    private showMapContextMenu(
        mapElement: HTMLElement,
        map: L.Map,
        regionName: string,
        event: L.LeafletMouseEvent
    ): void {
        this.hideMapContextMenu();

        const menu = document.createElement("div");
        menu.className = "map-context-menu";
        menu.style.left = `${event.containerPoint.x}px`;
        menu.style.top = `${event.containerPoint.y}px`;

        [
            { label: "Ver Colegios", mode: "colegios" as MapDrillMode }
        ].forEach((option: { label: string; mode: MapDrillMode }) => {
            const button = document.createElement("button");
            button.type = "button";
            button.textContent = option.label;
            button.onclick = () => {
                this.hideMapContextMenu();
                if (option.mode === "colegios") {
                    this.openSchoolListModal(regionName);
                    return;
                }

                this.showRegionDrillLayer(map, regionName, option.mode);
            };
            menu.appendChild(button);
        });

        mapElement.parentElement?.appendChild(menu);
        this.mapContextMenu = menu;
    }

    private hideMapContextMenu(): void {
        if (this.mapContextMenu) {
            this.mapContextMenu.remove();
            this.mapContextMenu = null;
        }
    }

    private showRegionDrillLayer(map: L.Map, regionName: string, mode: MapDrillMode): void {
        const layerKey = `${this.lastDatasetSignature}|${this.normalizeRegionForMap(regionName)}|${mode}|${map.getZoom()}`;
        if (this.activeMapDrillKey === layerKey && this.activeMapDrillLayer) {
            return;
        }

        if (this.activeMapDrillLayer) {
            this.activeMapDrillLayer.remove();
            this.activeMapDrillLayer = null;
        }

        const points = mode === "colegios"
            ? this.buildSchoolClusters(map, regionName)
            : this.buildMapPointSummaries(regionName, mode);
        const layer = L.layerGroup();

        points.forEach((point: MapPointSummary | SchoolCluster) => {
            const marker = this.createDrillMarker(point, mode);
            marker.addTo(layer);
        });

        layer.addTo(map);
        this.activeMapDrillLayer = layer;
        this.activeMapDrillKey = layerKey;

        if (points.length) {
            const bounds = L.latLngBounds(points.map((point: MapPointSummary | SchoolCluster) => [point.lat, point.lng] as L.LatLngTuple));
            map.fitBounds(bounds, { padding: [44, 44], maxZoom: mode === "colegios" ? 10 : 8 });
        }
    }

    private buildMapPointSummaries(regionName: string, mode: Exclude<MapDrillMode, "colegios">): MapPointSummary[] {
        const source = mode === "provincias"
            ? this.streamingIndexes.provinciaPoints
            : this.streamingIndexes.distritoPoints;

        return Array.from(source.values())
            .filter((point: MapPointAccumulator) => this.matchesRegion(point.regionKey, regionName))
            .map((point: MapPointAccumulator) => ({
                key: point.key,
                label: point.label,
                count: point.count,
                lat: point.latTotal / point.count,
                lng: point.lngTotal / point.count,
                riskBreakdown: point.riskBreakdown
            }))
            .sort((left: MapPointSummary, right: MapPointSummary) => right.count - left.count);
    }

    private buildSchoolClusters(map: L.Map, regionName: string): SchoolCluster[] {
        const zoom = map.getZoom();
        const cellSize = Math.max(42, 92 - zoom * 4);
        const clusters = new Map<string, SchoolCluster>();

        this.streamingIndexes.schoolPoints.forEach((point: SchoolPointIndex) => {
            if (!this.matchesRegion(point.regionKey, regionName)) {
                return;
            }

            const projected = map.project([point.lat, point.lng], zoom);
            const cellX = Math.floor(projected.x / cellSize);
            const cellY = Math.floor(projected.y / cellSize);
            const key = `${cellX}:${cellY}`;
            const cluster = clusters.get(key) || {
                key,
                count: 0,
                lat: 0,
                lng: 0,
                labels: [],
                riskBreakdown: { bajo: 0, medio: 0, alto: 0, muyAlto: 0 }
            };

            cluster.count += 1;
            cluster.lat += point.lat;
            cluster.lng += point.lng;
            if (cluster.labels.length < 6) {
                cluster.labels.push(point.label);
            }
            this.addRiskToBreakdown(cluster.riskBreakdown, point.riskLevel);
            clusters.set(key, cluster);
        });

        return Array.from(clusters.values())
            .map((cluster: SchoolCluster) => ({
                ...cluster,
                lat: cluster.lat / cluster.count,
                lng: cluster.lng / cluster.count
            }))
            .sort((left: SchoolCluster, right: SchoolCluster) => right.count - left.count);
    }

    private createDrillMarker(point: MapPointSummary | SchoolCluster, mode: MapDrillMode): L.Marker {
        const isCluster = mode === "colegios" && point.count > 1;
        const icon = L.divIcon({
            className: `map-drill-marker ${isCluster ? "map-school-cluster" : "map-drill-point"}`,
            html: `<span>${this.formatInteger(point.count)}</span>`,
            iconSize: isCluster ? [34, 34] : [28, 28],
            iconAnchor: isCluster ? [17, 17] : [14, 14]
        });
        const marker = L.marker([point.lat, point.lng], { icon });
        marker.bindTooltip(this.createDrillTooltip(point, mode), {
            direction: "top",
            opacity: 0.95,
            className: "region-leaflet-tooltip"
        });

        return marker;
    }

    private createDrillTooltip(point: MapPointSummary | SchoolCluster, mode: MapDrillMode): HTMLElement {
        const tooltip = document.createElement("div");
        const title = document.createElement("strong");
        title.textContent = "label" in point
            ? point.label
            : `Cluster de colegios (${this.formatInteger(point.count)})`;
        tooltip.appendChild(title);

        const rows: string[][] = [
            [mode === "colegios" ? "Colegios agrupados" : "Registros", this.formatInteger(point.count)],
            ["Riesgo Bajo", this.formatInteger(point.riskBreakdown.bajo)],
            ["Riesgo Medio", this.formatInteger(point.riskBreakdown.medio)],
            ["Riesgo Alto", this.formatInteger(point.riskBreakdown.alto)],
            ["Riesgo Muy Alto", this.formatInteger(point.riskBreakdown.muyAlto)]
        ];

        if (!("label" in point) && point.labels.length) {
            rows.push(["Muestra", point.labels.join(", ")]);
        }

        rows.forEach(([label, value]: string[]) => this.appendTooltipRow(tooltip, label, value));
        return tooltip;
    }

    private getRegionLayerStyle(
        feature: Feature<Geometry> | undefined,
        aggregateMap: Map<string, RegionAggregate>,
        mapHeatScale: MapHeatScale
    ): L.PathOptions {
        const regionName = feature ? this.getGeoJsonRegionName(feature) : "";
        const aggregate = this.findRegionAggregate(regionName, aggregateMap);
        const selected = this.isSelectedRegion(regionName);
        const fillColor = this.getRegionFillColor(aggregate, mapHeatScale);

        return {
            fillColor,
            fillOpacity: this.selectedRiskView === "Total"
                ? (selected ? 0.48 : neutralTotalRegionFillOpacity)
                : (selected ? 0.82 : 0.68),
            weight: selected ? 4 : 1.8,
            color: selected ? "#0f172a" : "#ffffff",
            opacity: 1
        };
    }

    private buildMapHeatScale(aggregates: RegionAggregate[]): MapHeatScale {
        return {
            maxTotalColegios: aggregates.reduce((max: number, aggregate: RegionAggregate) => (
                Math.max(max, aggregate.totalColegios)
            ), 0),
            riskHeatScores: this.buildRiskHeatScores(aggregates)
        };
    }

    private getRegionFillColor(aggregate: RegionAggregate | undefined, mapHeatScale: MapHeatScale): string {
        if (!aggregate) {
            return getRiskGradientColor(-1);
        }

        if (this.selectedRiskView === "Total") {
            return neutralTotalRegionFill;
        }

        const percentage = this.getSelectedRiskPercentage(aggregate);
        if (percentage <= 0) {
            return getRiskGradientColor(-1);
        }

        const score = mapHeatScale.riskHeatScores.get(this.normalizeRegionForMap(aggregate.region)) ?? -1;
        return getRiskGradientColor(score);
    }

    private buildRiskHeatScores(aggregates: RegionAggregate[]): Map<string, number> {
        const scores = new Map<string, number>();

        if (this.selectedRiskView === "Total") {
            return scores;
        }

        const positivePercentages = aggregates
            .map((aggregate: RegionAggregate) => ({
                key: this.normalizeRegionForMap(aggregate.region),
                percentage: this.getSelectedRiskPercentage(aggregate)
            }))
            .filter((item: { key: string; percentage: number }) => item.percentage > 0)
            .sort((left: { key: string; percentage: number }, right: { key: string; percentage: number }) => (
                left.percentage - right.percentage
            ));

        const rankedPercentages = positivePercentages.slice().reverse();
        const selectedPercentages = this.mapRankingLimit > 0
            ? rankedPercentages.slice(0, this.mapRankingLimit)
            : rankedPercentages;
        const total = selectedPercentages.length;

        selectedPercentages.forEach((item: { key: string; percentage: number }, index: number) => {
            const score = total <= 1 ? 1 : 1 - (index / (total - 1));
            scores.set(item.key, score);
        });

        return scores;
    }

    private getSelectedRiskPercentage(aggregate: RegionAggregate | undefined): number {
        if (!aggregate || this.selectedRiskView === "Total" || !aggregate.totalColegios) {
            return 0;
        }

        return this.getRiskValue(aggregate.riskBreakdown, this.selectedRiskView) / aggregate.totalColegios;
    }

    private getGeoJsonRegionName(feature: Feature<Geometry>): string {
        const properties = feature.properties || {};
        return String(
            properties.REGION ||
            properties.DEPARTAMEN ||
            properties.NOMBDEP ||
            properties.name ||
            ""
        );
    }

    private ensureOpenStreetMapTileLayer(map: L.Map): L.TileLayer {
        return L.tileLayer("https://tile.openstreetmap.org/{z}/{x}/{y}.png", {
            attribution: "OpenStreetMap",
            maxZoom: 18,
            minZoom: 4,
            crossOrigin: true
        }).addTo(map);
    }

    private createRegionTooltip(regionName: string, aggregate: RegionAggregate | undefined): HTMLElement {
        const tooltip = document.createElement("div");
        tooltip.className = "region-tooltip-content";
        const title = document.createElement("strong");
        title.className = "region-tooltip-title";
        title.textContent = aggregate?.region || regionName;
        tooltip.appendChild(title);

        const columns = document.createElement("div");
        columns.className = "region-tooltip-columns";

        const leftColumn = document.createElement("div");
        leftColumn.className = "region-tooltip-column";

        const leftTitle = document.createElement("strong");
        leftTitle.className = "region-tooltip-column-title";
        leftTitle.innerHTML = "Nivel de Riesgo<br/>por I.E.";
        leftColumn.appendChild(leftTitle);

        this.getRegionTooltipRows(aggregate).forEach(([label, value, className]: string[]) => {
            this.appendTooltipRow(leftColumn, label, value, className);
        });
        columns.appendChild(leftColumn);

        const regionAnalytics = this.getRegionAnalyticsEngine(regionName);
        const rightColumn = document.createElement("div");
        rightColumn.className = "region-tooltip-column";

        const unitsTitle = document.createElement("strong");
        unitsTitle.className = "region-tooltip-column-title";
        unitsTitle.innerHTML = "Intervenciones<br/>PRONIED";
        rightColumn.appendChild(unitsTitle);

        if (regionAnalytics) {
            const unitOrder: UnitKpiName[] = ["UGRD", "UGM", "UGEO", "UGME", "UGSC"];
            const unitMap = new Map(regionAnalytics.getUnitKpis().map((kpi: UnitKpiSummary) => [kpi.unit, kpi]));

            unitOrder.forEach((unit: UnitKpiName) => {
                const kpi = unitMap.get(unit);
                if (kpi) {
                    this.appendTooltipUnitRow(rightColumn, unit, this.formatInteger(kpi.colegios));
                }
            });
        }

        columns.appendChild(rightColumn);
        tooltip.appendChild(columns);

        return tooltip;
    }

    private getRegionTooltipRows(aggregate: RegionAggregate | undefined): string[][] {
        const totalColegios = aggregate?.totalColegios || 0;

        if (this.selectedRiskView !== "Total") {
            const selectedRiskCount = aggregate
                ? this.getRiskValue(aggregate.riskBreakdown, this.selectedRiskView)
                : 0;

            return [
                ["N° Colegios Totales", this.formatInteger(totalColegios), "region-tooltip-row-main"],
                [`Riesgo ${this.selectedRiskView}`, this.formatInteger(selectedRiskCount), this.getRiskTooltipClass(this.selectedRiskView)],
                ["Porcentaje", this.formatPercentFromRatio(this.getSelectedRiskPercentage(aggregate)), "region-tooltip-percent"]
            ];
        }

        const riskCount = (aggregate?.riskBreakdown.bajo || 0)
            + (aggregate?.riskBreakdown.medio || 0)
            + (aggregate?.riskBreakdown.alto || 0)
            + (aggregate?.riskBreakdown.muyAlto || 0);
        const unclassified = Math.max(0, totalColegios - riskCount);

        return [
            ["N° Colegios Totales", this.formatInteger(totalColegios), "region-tooltip-row-main"],
            ["N° Provincias", this.formatInteger(aggregate?.totalProvincias || 0), "region-tooltip-row-main"],
            ["Riesgo Bajo", this.formatInteger(aggregate?.riskBreakdown.bajo || 0), "risk-low"],
            ["Riesgo Medio", this.formatInteger(aggregate?.riskBreakdown.medio || 0), "risk-medium"],
            ["Riesgo Alto", this.formatInteger(aggregate?.riskBreakdown.alto || 0), "risk-high"],
            ["Riesgo Muy Alto", this.formatInteger(aggregate?.riskBreakdown.muyAlto || 0), "risk-critical"],
            ["Sin Clasificación", this.formatInteger(unclassified), "risk-unknown"]
        ];
    }

    private getRiskTooltipClass(risk: RiskLevel): string {
        if (risk === "Bajo") {
            return "risk-low";
        }

        if (risk === "Medio") {
            return "risk-medium";
        }

        if (risk === "Alto") {
            return "risk-high";
        }

        return "risk-critical";
    }

    private getRegionAnalyticsEngine(regionName: string): AnalyticsEngine | null {
        const regionKey = this.normalizeRegionForMap(regionName);
        return this.streamingIndexes.regionAnalytics.get(regionKey) || null;
    }

    private appendTooltipRow(tooltip: HTMLElement, label: string, value: string, className = ""): void {
        const row = document.createElement("div");
        row.className = `region-tooltip-row ${className}`.trim();
        const labelElement = document.createElement("span");
        labelElement.textContent = label;
        row.appendChild(labelElement);

        const valueElement = document.createElement("b");
        valueElement.textContent = value;
        row.appendChild(valueElement);
        tooltip.appendChild(row);
    }

    private appendTooltipUnitRow(tooltip: HTMLElement, unit: UnitKpiName, value: string): void {
        const row = document.createElement("div");
        row.className = "region-tooltip-row";

        const labelWrapper = document.createElement("span");
        labelWrapper.className = "region-tooltip-unit-label";

        const iconWrapper = document.createElement("span");
        iconWrapper.className = `region-tooltip-unit-icon unit-icon-${unit.toLowerCase()}`;
        this.appendIconSvg(iconWrapper, this.getUnitTheme(unit).icon);
        labelWrapper.appendChild(iconWrapper);

        const text = document.createElement("span");
        text.className = "region-tooltip-unit-text";
        text.textContent = unit;
        labelWrapper.appendChild(text);

        row.appendChild(labelWrapper);

        const valueElement = document.createElement("b");
        valueElement.textContent = value;
        row.appendChild(valueElement);
        tooltip.appendChild(row);
    }

    private removeLeafletMap(): void {
        this.hideMapContextMenu();
        if (this.activeMapDrillLayer) {
            this.activeMapDrillLayer.remove();
            this.activeMapDrillLayer = null;
            this.activeMapDrillKey = "";
        }

        if (this.leafletMap) {
            this.leafletMap.remove();
            this.leafletMap = null;
        }
    }

    private appendMapResetButton(map: HTMLElement): void {
        const button = document.createElement("button");
        button.className = "map-reset-button";
        button.type = "button";
        button.textContent = "R";
        button.title = "Volver al mapa completo";
        button.onclick = () => {
            this.resetMapFocus(true);
            this.renderCurrentView();
        };
        map.appendChild(button);
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

    private appendMapRankingControl(map: HTMLElement, engine: AnalyticsEngine): void {
        if (this.selectedRiskView === "Total") {
            return;
        }

        const panel = document.createElement("aside");
        panel.className = "map-ranking-card";

        const header = document.createElement("div");
        header.className = "map-ranking-header";

        const title = document.createElement("strong");
        title.textContent = "Ranking";
        header.appendChild(title);

        const input = document.createElement("input");
        input.type = "number";
        input.min = "0";
        input.max = "26";
        input.step = "1";
        input.value = this.mapRankingLimit.toString();
        input.title = "Cantidad de regiones a colorear";
        input.onchange = () => {
            const nextValue = Math.max(0, Math.min(26, Math.floor(Number(input.value) || 0)));
            this.mapRankingLimit = nextValue;
            this.renderCurrentView();
        };
        header.appendChild(input);
        panel.appendChild(header);

        const hint = document.createElement("small");
        hint.textContent = `% Riesgo ${this.selectedRiskView}`;
        panel.appendChild(hint);

        const list = document.createElement("ol");
        list.className = "map-ranking-list";
        const rankingRows = this.getMapRankingRows(engine);
        panel.style.setProperty("--ranking-list-height", `${Math.max(1, rankingRows.length) * 24}px`);
        rankingRows.forEach((row: { region: string; percentage: number; score: number }) => {
            const item = document.createElement("li");

            const swatch = document.createElement("span");
            swatch.style.background = getRiskGradientColor(row.score);
            item.appendChild(swatch);

            const label = document.createElement("b");
            label.textContent = row.region;
            label.title = row.region;
            item.appendChild(label);

            const value = document.createElement("em");
            value.textContent = this.formatRoundedPercentFromRatio(row.percentage);
            item.appendChild(value);
            list.appendChild(item);
        });

        panel.appendChild(list);
        map.appendChild(panel);
    }

    private getMapRankingRows(engine: AnalyticsEngine): { region: string; percentage: number; score: number }[] {
        if (this.selectedRiskView === "Total") {
            return [];
        }

        const rows = this.getRegionAggregates(engine)
            .map((aggregate: RegionAggregate) => ({
                region: aggregate.region,
                percentage: this.getSelectedRiskPercentage(aggregate)
            }))
            .filter((row: { region: string; percentage: number }) => row.percentage > 0)
            .sort((left: { region: string; percentage: number }, right: { region: string; percentage: number }) => (
                right.percentage - left.percentage
            ));
        const selectedRows = this.mapRankingLimit > 0 ? rows.slice(0, this.mapRankingLimit) : rows;
        const total = selectedRows.length;

        return selectedRows.map((row: { region: string; percentage: number }, index: number) => ({
            ...row,
            score: total <= 1 ? 1 : 1 - (index / (total - 1))
        }));
    }

    private appendRiskToggle(map: HTMLElement): void {
        const toggle = document.createElement("div");
        toggle.className = "map-risk-toggle";

        const label = document.createElement("span");
        label.textContent = "Vista por Riesgo:";
        toggle.appendChild(label);

        riskMapViews.forEach((mode: RiskMapView) => {
            const button = document.createElement("button");
            button.className = `map-risk-toggle-button${mode === this.selectedRiskView ? " map-risk-toggle-button-active" : ""}`;
            button.type = "button";
            button.textContent = mode;
            button.onclick = () => {
                this.selectedRiskView = mode;
                this.detailPage = 1;
                this.internalFilterAnalyticsCache = new Map<string, AnalyticsEngine>();
                this.selectedRegionAnalyticsCache = null;
                this.selectedRegionAnalyticsCacheKey = "";
                this.renderCurrentView();
            };
            toggle.appendChild(button);
        });

        map.appendChild(toggle);
    }

    private getSelectedRegionEngine(engine: AnalyticsEngine): AnalyticsEngine {
        if (!this.selectedRegion) {
            return engine;
        }

        const cacheKey = `${this.getInternalFilterCacheKey()}|${this.normalizeRegionForMap(this.selectedRegion)}`;
        if (this.enableSelectedRegionCache && this.selectedRegionAnalyticsCache && this.selectedRegionAnalyticsCacheKey === cacheKey) {
            return this.selectedRegionAnalyticsCache;
        }

        const filteredRows = this.accumulatedRows.filter((record: SourceRecord) => (
            this.recordMatchesDashboardFilters(record) &&
            this.isRegionNameSelected(record.region)
        ));
        const filteredEngine = filteredRows.length
            ? AnalyticsEngine.build(filteredRows)
            : AnalyticsEngine.createEmpty();
        if (!filteredRows.length) {
            filteredEngine.finalizeBuild();
        }

        if (this.enableSelectedRegionCache) {
            this.selectedRegionAnalyticsCache = filteredEngine;
            this.selectedRegionAnalyticsCacheKey = cacheKey;
        }

        return filteredEngine;
    }

    private isRecordInSelectedRegion(record: SourceRecord): boolean {
        if (!this.selectedRegion) {
            return true;
        }

        return this.isRegionNameSelected(record.region);
    }

    private isSelectedRegion(regionName: string): boolean {
        const focusedRegion = this.getFocusedRegion();
        if (!focusedRegion) {
            return false;
        }

        return this.matchesRegion(regionName, focusedRegion);
    }

    private isRegionNameSelected(regionName: string): boolean {
        if (!this.selectedRegion) {
            return false;
        }

        return this.matchesRegion(regionName, this.selectedRegion);
    }

    private getFocusedRegion(): string | null {
        return this.selectedRegion || this.autoFocusedRegion;
    }

    private matchesRegion(regionName: string, selectedRegionName: string): boolean {
        const selectedRegion = this.normalizeRegionForMap(selectedRegionName);
        const region = this.normalizeRegionForMap(regionName);

        return region === selectedRegion || this.getRegionAliases(selectedRegion).includes(region);
    }

    private getRegionAggregates(engine: AnalyticsEngine): RegionAggregate[] {
        return engine.getRegions().map((summary: BucketSummary) => ({
            region: summary.label,
            totalColegios: summary.colegiosUnicos,
            totalProvincias: summary.provinciasUnicas,
            montoInversion: summary.montoTotal,
            estudiantesBeneficiarios: summary.beneficiariosTotal,
            numeroSolicitudes: summary.solicitudesUnicas,
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

    private addRiskToBreakdown(riskBreakdown: RiskBuckets, value: string): void {
        const normalized = this.normalizeText(value);

        if (normalized === "BAJO") {
            riskBreakdown.bajo += 1;
        } else if (normalized === "MEDIO") {
            riskBreakdown.medio += 1;
        } else if (normalized === "ALTO") {
            riskBreakdown.alto += 1;
        } else {
            riskBreakdown.muyAlto += 1;
        }
    }

    private hasValidCoordinates(record: SourceRecord): boolean {
        return Number.isFinite(record.latitud) &&
            Number.isFinite(record.longitud) &&
            record.latitud >= -20 &&
            record.latitud <= 2 &&
            record.longitud >= -82 &&
            record.longitud <= -68;
    }

    private shortRegionLabel(region: string): string {
        return region
            .replace("LIMA METROPOLITANA", "LIMA MET.")
            .replace("LIMA PROVINCIAS", "LIMA PROV.");
    }

    private appendUgmeView(engine: AnalyticsEngine, mapEngine: AnalyticsEngine = engine): void {
        const view = document.createElement("section");
        view.className = "dashboard-root ugme-dashboard-view with-internal-filters";
        this.appendInternalFilterBar(view);

        const layout = document.createElement("div");
        layout.className = "main-dashboard-layout";

        const mapPanel = document.createElement("section");
        mapPanel.className = "map-panel";
        this.appendPeruRiskMap(mapPanel, mapEngine);
        layout.appendChild(mapPanel);

        const unitsPanel = this.createUgmeUnitsPanel(engine);
        layout.appendChild(unitsPanel);
        view.appendChild(layout);

        this.target.appendChild(view);
    }

    private createUgmeUnitsPanel(engine: AnalyticsEngine): HTMLElement {
        const unitEngine = this.getSelectedRegionEngine(engine);
        const unitsPanel = document.createElement("section");
        unitsPanel.className = "units-panel ugme-units-panel";

        this.appendInformationBanner(unitsPanel, unitEngine);

        const title = document.createElement("h1");
        title.textContent = this.selectedRegion
            ? `UGME - ${this.selectedRegion}`
            : "UGME";
        unitsPanel.appendChild(title);

        const kpi = unitEngine.getUnitKpi("UGME");
        const card = this.createUnitCard(kpi, true);
        card.classList.add("ugme-card");
        unitsPanel.appendChild(card);

        const backButton = document.createElement("button");
        backButton.className = "floating-next-button floating-back-button";
        backButton.type = "button";
        backButton.title = "Volver a la pantalla principal";
        backButton.setAttribute("aria-label", "Volver a la pantalla principal");
        backButton.textContent = "<";
        backButton.onclick = () => this.switchVisualView("summary");
        unitsPanel.appendChild(backButton);

        return unitsPanel;
    }

    private appendInformationBanner(container: HTMLElement, engine: AnalyticsEngine): void {
        const totals = engine.getTotals();
        const scope = this.getInformationScope(engine);
        const banner = document.createElement("section");
        banner.className = `information-banner${scope.isSingleSchool ? " information-banner-school" : ""}`;

        if (scope.header) {
            const header = document.createElement("div");
            header.className = "information-banner-header";

            const title = document.createElement("h2");
            title.textContent = scope.header;
            header.appendChild(title);

            const info = document.createElement("span");
            info.textContent = "i";
            header.appendChild(info);
            banner.appendChild(header);
        }

        const scopeTitle = document.createElement("h3");
        scopeTitle.textContent = scope.scopeTitle;
        banner.appendChild(scopeTitle);

        const metrics = document.createElement("div");
        metrics.className = "information-metric-grid";
        metrics.appendChild(this.createInformationMetric("IE", this.formatInteger(totals.colegiosUnicos), "Instituciones Educativas"));
        metrics.appendChild(this.createInformationMetric("EB", this.formatInteger(totals.beneficiariosTotal), "Estudiantes Beneficiarios"));
        banner.appendChild(metrics);

        container.appendChild(banner);
    }

    private createInformationMetric(icon: string, value: string, label: string): HTMLElement {
        const card = document.createElement("article");
        card.className = "information-metric-card";

        const iconElement = document.createElement("span");
        iconElement.textContent = icon;
        card.appendChild(iconElement);

        const content = document.createElement("div");
        const strong = document.createElement("strong");
        strong.textContent = value;
        content.appendChild(strong);

        const small = document.createElement("small");
        small.textContent = label;
        content.appendChild(small);
        card.appendChild(content);

        return card;
    }

    private getInformationScope(engine: AnalyticsEngine): { header: string; scopeTitle: string; isSingleSchool: boolean } {
        const singleSchool = this.getSingleSchoolDetail(engine);
        if (singleSchool) {
            return {
                header: singleSchool.nombreLocal && singleSchool.nombreLocal !== "-"
                    ? singleSchool.nombreLocal
                    : singleSchool.codigoLocal,
                scopeTitle: [singleSchool.region, singleSchool.provincia, singleSchool.distrito]
                    .filter((value: string) => !!value && value !== "-")
                    .join(" / ") || singleSchool.codigoLocal,
                isSingleSchool: true
            };
        }

        const regions = engine.getRegions();
        if (!regions.length || regions.length > 3) {
            return {
                header: "",
                scopeTitle: this.selectedRegion || "A NIVEL NACIONAL",
                isSingleSchool: false
            };
        }

        if (regions.length === 1) {
            const region = regions[0];
            const provincias = engine.getProvincias(region.key);
            const scopeParts = [region.label];

            if (provincias.length === 1) {
                const provincia = provincias[0];
                const distritos = engine.getDistritos(provincia.key);
                scopeParts.push(provincia.label);

                if (distritos.length === 1) {
                    scopeParts.push(distritos[0].label);
                }
            }

            return {
                header: "",
                scopeTitle: scopeParts.join(" / "),
                isSingleSchool: false
            };
        }

        return {
            header: "",
            scopeTitle: regions.map((region: BucketSummary) => region.label).join(", "),
            isSingleSchool: false
        };
    }

    private getSingleSchoolDetail(engine: AnalyticsEngine): UnitDetailRow | null {
        if (engine.getTotals().colegiosUnicos !== 1) {
            return null;
        }

        for (const unitRows of this.streamingIndexes.detailRowsByUnit.values()) {
            for (const row of unitRows.values()) {
                if (engine.getBucketByColegio(row.codigoLocal) && (!this.selectedRegion || this.isRegionNameSelected(row.region))) {
                    return row;
                }
            }
        }

        return null;
    }

    private createUnitCard(kpi: UnitKpiSummary, showUgmeDetail: boolean): HTMLElement {
        if (showUgmeDetail && kpi.unit === "UGME") {
            return this.createUgmeDetailCard(kpi);
        }

        const theme = this.getUnitTheme(kpi.unit);
        const card = document.createElement("section");
        card.className = `unit-card unit-card-${kpi.unit.toLowerCase()}`;
        card.style.setProperty("--unit-color", theme.color);
        card.style.setProperty("--unit-bg", theme.background);

        const header = document.createElement("div");
        header.className = "unit-card-header";

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

    private createUgmeDetailCard(kpi: UnitKpiSummary): HTMLElement {
        const theme = this.getUnitTheme("UGME");
        const card = document.createElement("section");
        card.className = "ugme-detail-card";
        card.style.setProperty("--unit-color", theme.color);
        card.style.setProperty("--unit-bg", theme.background);

        const hero = document.createElement("section");
        hero.className = "ugme-hero";

        const header = document.createElement("div");
        header.className = "ugme-hero-header";

        const icon = document.createElement("div");
        icon.className = "ugme-hero-icon";
        icon.textContent = "ME";
        header.appendChild(icon);

        const titleGroup = document.createElement("div");
        const title = document.createElement("h2");
        title.textContent = "UGME";
        titleGroup.appendChild(title);

        const subtitle = document.createElement("p");
        subtitle.textContent = theme.name;
        titleGroup.appendChild(subtitle);
        header.appendChild(titleGroup);

        hero.appendChild(header);

        const kpiGrid = document.createElement("div");
        kpiGrid.className = "ugme-hero-kpis";
        [
            { icon: "map-pin", label: "N° Regiones", value: this.formatInteger(kpi.regiones) },
            { icon: "school", label: "N° Colegios", value: this.formatInteger(kpi.colegios) },
            { icon: "money", label: "Monto Total", value: this.formatMillions(kpi.montoIntervencion) }
        ].forEach((item: UnitKpiCell) => kpiGrid.appendChild(this.createUgmeHeroKpi(item)));
        hero.appendChild(kpiGrid);
        card.appendChild(hero);

        const sections = document.createElement("div");
        sections.className = "ugme-detail-sections";
        this.appendUgmeGroupSection(sections, "MODULOS PREFABRICADOS", this.getUgmeModularGroups(kpi.grupos), "ugme-section-modulares");
        this.appendUgmeGroupSection(sections, "MOBILIARIO Y EQUIPAMIENTO", this.getUgmeMobiliarioGroups(kpi.grupos), "ugme-section-mobiliario");
        card.appendChild(sections);

        return card;
    }

    private createUgmeHeroKpi(item: UnitKpiCell): HTMLElement {
        const card = document.createElement("article");
        card.className = "ugme-hero-kpi";

        const icon = document.createElement("span");
        this.appendIconSvg(icon, item.icon);
        card.appendChild(icon);

        const content = document.createElement("div");
        const value = document.createElement("strong");
        value.textContent = item.value;
        content.appendChild(value);

        const label = document.createElement("small");
        label.textContent = item.label;
        content.appendChild(label);
        card.appendChild(content);

        return card;
    }

    private getUnitTheme(unit: UnitKpiName): UnitTheme {
        const themes: Record<UnitKpiName, UnitTheme> = {
            UGRD: {
                color: "#f97316",
                background: "#fff7ed",
                icon: "shield",
                name: "Unidad Gerencial de Reconstrucción frente a Desastres"
            },
            UGSC: {
                color: "#7c3aed",
                background: "#f5f3ff",
                icon: "briefcase",
                name: "Unidad Gerencial de Supervision y Convenios"
            },
            UGEO: {
                color: "#2563eb",
                background: "#eff6ff",
                icon: "home",
                name: "Unidad Gerencial de Estudios y Obras"
            },
            UGM: {
                color: "#16a34a",
                background: "#f0fdf4",
                icon: "home",
                name: "Unidad Gerencial de Mantenimiento"
            },
            UGME: {
                color: "#0891b2",
                background: "#ecfeff",
                icon: "package",
                name: "Unidad Gerencial de Mobiliario y Equipamiento"
            }
        };

        return themes[unit];
    }

    private getUnitKpiCells(kpi: UnitKpiSummary): UnitKpiCell[] {
        if (kpi.unit === "UGSC") {
            return [
                { icon: "clock", label: "N° Regiones", value: this.formatInteger(kpi.regiones) },
                { icon: "school", label: "N° Colegios", value: this.formatInteger(kpi.colegios) },
                { icon: "document", label: "N° Solicitudes", value: this.formatInteger(kpi.solicitudes) },
                { icon: "users", label: "Estudiantes Beneficiarios", value: this.formatInteger(kpi.beneficiarios) },
                { icon: "document", label: "Solicitudes en revision", value: this.formatInteger(kpi.solicitudesRevision) },
                { icon: "document", label: "Solicitudes culminadas", value: this.formatInteger(kpi.solicitudesCulminadas) },
                { icon: "money", label: "Monto de Inversion", value: this.formatNumber(kpi.montoIntervencion) }
            ];
        }

        if (kpi.unit === "UGM") {
            const porcentajeTransferencias = kpi.montoAsignado
                ? (kpi.montoTransferencia / kpi.montoAsignado) * 100
                : 0;

            return [
                { icon: "clock", label: "N° Regiones", value: this.formatInteger(kpi.regiones) },
                { icon: "school", label: "N° Colegios", value: this.formatInteger(kpi.colegios) },
                { icon: "money", label: "Monto Programado", value: this.formatNumber(kpi.montoAsignado) },
                { icon: "money", label: "Monto Transferido", value: this.formatNumber(kpi.montoTransferencia) },
                { icon: "percent", label: "% Transferencias", value: this.formatPercent(porcentajeTransferencias) },
                { icon: "money", label: "Monto Retirado", value: this.formatNumber(kpi.montoRetirado) },
                { icon: "school", label: "N° Colegios con Transferencia", value: this.formatInteger(kpi.colegiosConTransferencia) },
                { icon: "school", label: "N° Colegios con Retiro", value: this.formatInteger(kpi.colegiosConRetiro) }
            ];
        }

        if (kpi.unit === "UGME") {
            return [
                { icon: "clock", label: "N° Regiones", value: this.formatInteger(kpi.regiones) },
                { icon: "school", label: "N° Colegios", value: this.formatInteger(kpi.colegios) },
                { icon: "money", label: "Monto", value: this.formatMillions(kpi.montoIntervencion) }
            ];
        }

        return [
            { icon: "clock", label: "N° Regiones", value: this.formatInteger(kpi.regiones) },
            { icon: "school", label: "N° Colegios", value: this.formatInteger(kpi.colegios) },
            { icon: "money", label: "Monto de inversion", value: this.formatNumber(kpi.montoIntervencion) },
            { icon: "users", label: "Estudiantes beneficiarios", value: this.formatInteger(kpi.beneficiarios) }
        ];
    }

    private createKpiCell(item: UnitKpiCell): HTMLElement {
        const cell = document.createElement("div");
        cell.className = "unit-kpi-cell";

        const icon = document.createElement("span");
        icon.className = "unit-kpi-icon";
        this.appendIconSvg(icon, item.icon);
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

    private appendIconSvg(container: HTMLElement, iconName: string): void {
        const svgNamespace = "http://www.w3.org/2000/svg";
        const svg = document.createElementNS(svgNamespace, "svg");
        svg.setAttribute("viewBox", "0 0 24 24");
        svg.setAttribute("aria-hidden", "true");
        svg.setAttribute("focusable", "false");

        this.getIconShapes(iconName).forEach((shape: SvgIconShape) => {
            const element = document.createElementNS(svgNamespace, shape.tag);
            Object.entries(shape.attrs).forEach(([key, value]: [string, string]) => element.setAttribute(key, value));
            svg.appendChild(element);
        });

        container.replaceChildren(svg);
    }

    private getIconShapes(iconName: string): SvgIconShape[] {
        const path = (d: string): SvgIconShape => ({ tag: "path", attrs: { d } });
        const circle = (cx: string, cy: string, r: string): SvgIconShape => ({ tag: "circle", attrs: { cx, cy, r } });
        const rect = (x: string, y: string, width: string, height: string, rx: string): SvgIconShape => ({
            tag: "rect",
            attrs: { x, y, width, height, rx }
        });

        const icons: Record<string, SvgIconShape[]> = {
            bolt: [
                path("M13 2L4 14h7l-1 8 9-12h-7z")
            ],
            briefcase: [
                rect("5", "7", "14", "12", "2"),
                path("M9 7V5h6v2"),
                path("M9 12h6")
            ],
            building: [
                path("M5 21V6l7-3 7 3v15"),
                path("M8 10h2"),
                path("M14 10h2"),
                path("M8 14h2"),
                path("M14 14h2"),
                path("M10 21v-4h4v4")
            ],
            chair: [
                path("M7 12h10"),
                path("M8 12V7a4 4 0 0 1 8 0v5"),
                path("M6 12v7"),
                path("M18 12v7"),
                path("M7 16h10")
            ],
            clock: [
                circle("12", "12", "8"),
                path("M12 8v5l3 2")
            ],
            cube: [
                path("M12 3l8 4.5v9L12 21l-8-4.5v-9z"),
                path("M12 12l8-4.5"),
                path("M12 12L4 7.5"),
                path("M12 12v9")
            ],
            document: [
                path("M8 4h6l3 3v13H8z"),
                path("M14 4v4h4"),
                path("M10 12h5"),
                path("M10 16h5")
            ],
            flask: [
                path("M9 3h6"),
                path("M10 3v6l-5 8a3 3 0 0 0 2.6 4h8.8a3 3 0 0 0 2.6-4l-5-8V3"),
                path("M8 16h8")
            ],
            home: [
                path("M4 11l8-7 8 7"),
                path("M6 10v10h12V10"),
                path("M10 20v-6h4v6")
            ],
            "map-pin": [
                path("M12 21s7-5.2 7-11a7 7 0 0 0-14 0c0 5.8 7 11 7 11z"),
                circle("12", "10", "2.5")
            ],
            monitor: [
                rect("4", "5", "16", "11", "1.5"),
                path("M9 20h6"),
                path("M12 16v4"),
                path("M8 9l2-2"),
                path("M8 12l5-5")
            ],
            money: [
                circle("12", "12", "8"),
                path("M14.5 9.5c-.7-.6-1.6-.9-2.6-.8-1.2.1-2.1.8-2.1 1.8 0 1.1 1.1 1.5 2.5 1.8 1.5.4 2.4.8 2.4 1.9 0 1-.9 1.8-2.2 1.9-1.1.1-2.2-.2-3-.9"),
                path("M12 7v10")
            ],
            network: [
                circle("12", "5", "2"),
                circle("6", "18", "2"),
                circle("18", "18", "2"),
                path("M11.2 7L7 16"),
                path("M12.8 7L17 16"),
                path("M8 18h8")
            ],
            package: [
                path("M5 8l7-4 7 4-7 4z"),
                path("M5 8v8l7 4 7-4V8"),
                path("M12 12v8")
            ],
            percent: [
                path("M7 17L17 7"),
                circle("8", "8", "2"),
                circle("16", "16", "2")
            ],
            presentation: [
                rect("4", "5", "16", "11", "1.5"),
                path("M8 20l4-4 4 4"),
                path("M9 12l2-2 2 1 3-4")
            ],
            restroom: [
                circle("8", "5", "1.6"),
                circle("16", "5", "1.6"),
                path("M7 8h2l1 5H6z"),
                path("M15 8h2v10"),
                path("M15 13h3"),
                path("M7 13v6"),
                path("M9 13v6")
            ],
            school: [
                path("M4 20V9l8-5 8 5v11"),
                path("M8 20v-6h8v6"),
                path("M9 10h6")
            ],
            shield: [
                path("M12 21s7-3.5 7-10V5l-7-3-7 3v6c0 6.5 7 10 7 10z"),
                path("M12 7v7"),
                path("M9.5 10.5h5")
            ],
            tools: [
                path("M14 6l4-4 4 4-4 4z"),
                path("M2 22l8-8"),
                path("M16 14l6 6"),
                path("M18 20l2-2"),
                path("M3 6l5 5"),
                path("M5 4l5 5")
            ],
            users: [
                path("M8 19c0-2.2 1.8-4 4-4s4 1.8 4 4"),
                circle("12", "9", "3"),
                path("M17 11c1.6.3 3 1.7 3 3.5"),
                path("M7 11c-1.6.3-3 1.7-3 3.5")
            ]
        };

        return icons[iconName] || icons.document;
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
                ["Monto", this.formatMillions(grupo.montoIntervencion)]
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

    private appendUgmeGroupSection(container: HTMLElement, titleText: string, grupos: GroupMetricSummary[], className: string): void {
        const section = document.createElement("section");
        section.className = `ugme-detail-section ${className}`;

        const title = document.createElement("h3");
        title.textContent = titleText;
        section.appendChild(title);

        const header = document.createElement("div");
        header.className = "ugme-detail-table-header";

        const headerIcon = document.createElement("span");
        this.appendIconSvg(headerIcon, className === "ugme-section-modulares" ? "cube" : "chair");
        header.appendChild(headerIcon);

        ["Item", "Cantidad", "Monto"].forEach((text: string) => {
            const label = document.createElement("strong");
            label.textContent = text;
            header.appendChild(label);
        });
        section.appendChild(header);

        const grid = document.createElement("div");
        grid.className = "ugme-detail-group-grid";
        grupos.forEach((grupo: GroupMetricSummary) => grid.appendChild(this.createUgmeDetailGroupCard(grupo)));
        section.appendChild(grid);
        container.appendChild(section);
    }

    private createUgmeDetailGroupCard(grupo: GroupMetricSummary): HTMLElement {
        const card = document.createElement("article");
        const groupClass = this.getUgmeGroupClass(grupo.grupo);
        card.className = `ugme-detail-group-row-card ${groupClass}`;

        const icon = document.createElement("span");
        icon.className = "ugme-detail-item-icon";
        this.appendIconSvg(icon, this.getUgmeGroupIcon(grupo.grupo));
        card.appendChild(icon);

        const item = document.createElement("strong");
        item.className = "ugme-detail-item-name";
        item.textContent = grupo.grupo;
        card.appendChild(item);

        const quantity = document.createElement("strong");
        quantity.className = "ugme-detail-item-quantity";
        quantity.textContent = this.formatInteger(grupo.cantidad);
        card.appendChild(quantity);

        const amount = document.createElement("strong");
        amount.className = "ugme-detail-item-amount";
        amount.textContent = this.formatMillions(grupo.montoIntervencion);
        card.appendChild(amount);

        return card;
    }

    private getUgmeModularGroups(grupos: GroupMetricSummary[]): GroupMetricSummary[] {
        const modularNames = new Set<string>([
            "ESCUELA MODULAR",
            "AULA MODULAR",
            "KIT DE PARARRAYO",
            "MODULO SS.HH.",
            "REDES COMPLEMENTARIAS"
        ]);

        return grupos.filter((grupo: GroupMetricSummary) => modularNames.has(grupo.grupo));
    }

    private getUgmeMobiliarioGroups(grupos: GroupMetricSummary[]): GroupMetricSummary[] {
        const modularNames = new Set<string>([
            "ESCUELA MODULAR",
            "AULA MODULAR",
            "KIT DE PARARRAYO",
            "MODULO SS.HH.",
            "REDES COMPLEMENTARIAS"
        ]);

        return grupos.filter((grupo: GroupMetricSummary) => !modularNames.has(grupo.grupo));
    }

    private getUgmeGroupIcon(grupo: string): string {
        const icons: Record<string, string> = {
            "ESCUELA MODULAR": "building",
            "AULA MODULAR": "presentation",
            "KIT DE PARARRAYO": "bolt",
            "MODULO SS.HH.": "restroom",
            "REDES COMPLEMENTARIAS": "network",
            MOBILIARIO: "chair",
            EQUIPAMIENTO: "monitor",
            LABORATORIOS: "flask",
            TALLERES: "tools"
        };

        return icons[grupo] || "package";
    }

    private getUgmeGroupClass(grupo: string): string {
        const classes: Record<string, string> = {
            MOBILIARIO: "ugme-group-green",
            EQUIPAMIENTO: "ugme-group-teal",
            LABORATORIOS: "ugme-group-purple",
            TALLERES: "ugme-group-orange"
        };

        return classes[grupo] || "ugme-group-blue";
    }

    private openDetailModal(unit: UnitKpiName): void {
        this.detailModalUnit = unit;
        this.isDetailModalOpen = true;
        this.detailSearchText = "";
        this.detailPage = 1;
        this.refreshDetailModal();
    }

    private closeDetailModal(): void {
        this.isDetailModalOpen = false;
        this.detailModalUnit = null;
        this.detailSearchText = "";
        this.detailPage = 1;
        this.removeDetailModal();
    }

    private refreshDetailModal(): void {
        this.removeDetailModal();
        if (this.isDetailModalOpen && this.detailModalUnit) {
            this.appendDetailModal(this.detailModalUnit);
        }
    }

    private removeDetailModal(): void {
        this.target.querySelectorAll(".detail-modal-backdrop").forEach((modal: Element) => modal.remove());
    }

    private openSchoolListModal(regionFilter: string | null): void {
        this.isSchoolListModalOpen = true;
        this.schoolListRegionFilter = regionFilter;
        this.selectedSchoolKey = null;
        this.detailSearchText = "";
        this.detailPage = 1;
        this.refreshSchoolModals();
    }

    private closeSchoolListModal(): void {
        this.isSchoolListModalOpen = false;
        this.schoolListRegionFilter = null;
        this.selectedSchoolKey = null;
        this.detailSearchText = "";
        this.detailPage = 1;
        this.removeSchoolModals();
    }

    private openSchoolInterventionsModal(row: SchoolSummaryRow): void {
        this.selectedSchoolKey = this.normalizeText(row.codigoLocal);
        this.selectedSchoolTab = this.getFirstSchoolUnit(row) || "UGRD";
        this.refreshSchoolModals();
    }

    private closeSchoolInterventionsModal(): void {
        this.selectedSchoolKey = null;
        this.removeSchoolInterventionsModal();
    }

    private refreshSchoolModals(): void {
        this.removeSchoolModals();
        if (this.isSchoolListModalOpen) {
            this.appendSchoolListModal();
        }

        if (this.selectedSchoolKey) {
            this.appendSchoolInterventionsModal(this.selectedSchoolKey);
        }
    }

    private removeSchoolModals(): void {
        this.target.querySelectorAll(".school-list-modal-backdrop").forEach((modal: Element) => modal.remove());
        this.removeSchoolInterventionsModal();
    }

    private removeSchoolInterventionsModal(): void {
        this.target.querySelectorAll(".school-interventions-modal-backdrop").forEach((modal: Element) => modal.remove());
    }

    private appendSchoolListModal(): void {
        const backdrop = document.createElement("div");
        backdrop.className = "detail-modal-backdrop school-list-modal-backdrop";
        backdrop.onclick = (event: MouseEvent) => {
            if (event.target === backdrop) {
                this.closeSchoolListModal();
            }
        };

        const modal = document.createElement("section");
        modal.className = "detail-modal school-list-modal";

        const header = document.createElement("header");
        header.className = "detail-modal-header school-list-modal-header";
        const titleGroup = document.createElement("div");
        const title = document.createElement("h2");
        title.textContent = this.schoolListRegionFilter
            ? `Colegios - ${this.schoolListRegionFilter}`
            : "Colegios";
        titleGroup.appendChild(title);
        const subtitle = document.createElement("p");
        subtitle.textContent = "Seleccione un colegio para revisar sus intervenciones por unidad gerencial.";
        titleGroup.appendChild(subtitle);
        header.appendChild(titleGroup);

        const closeIcon = document.createElement("button");
        closeIcon.className = "detail-modal-close";
        closeIcon.type = "button";
        closeIcon.textContent = "X";
        closeIcon.onclick = () => this.closeSchoolListModal();
        header.appendChild(closeIcon);
        modal.appendChild(header);

        const allRows = this.getSchoolSummaryRows();
        const filteredRows = this.filterSchoolSummaryRows(allRows);
        const totalPages = Math.max(1, Math.ceil(filteredRows.length / this.detailPageSize));
        this.detailPage = Math.min(this.detailPage, totalPages);
        const pageStart = (this.detailPage - 1) * this.detailPageSize;
        const pageRows = filteredRows.slice(pageStart, pageStart + this.detailPageSize);

        const toolbar = document.createElement("div");
        toolbar.className = "detail-modal-toolbar";

        const search = document.createElement("input");
        search.className = "detail-search-input";
        search.type = "search";
        search.placeholder = "Buscar por codigo, colegio o ubicacion";
        search.value = this.detailSearchText;
        search.oninput = () => {
            this.detailSearchText = search.value;
            this.detailPage = 1;
            this.refreshSchoolModals();
        };
        toolbar.appendChild(search);

        const count = document.createElement("span");
        count.className = "detail-count";
        count.textContent = `Mostrando ${pageRows.length} de ${filteredRows.length} colegios`;
        toolbar.appendChild(count);
        modal.appendChild(toolbar);

        this.appendSchoolListTable(modal, pageRows);
        this.appendSchoolListPagination(modal, totalPages);

        const footer = document.createElement("footer");
        footer.className = "detail-modal-footer";
        const closeButton = document.createElement("button");
        closeButton.type = "button";
        closeButton.textContent = "Cerrar";
        closeButton.onclick = () => this.closeSchoolListModal();
        footer.appendChild(closeButton);
        modal.appendChild(footer);

        backdrop.appendChild(modal);
        this.target.appendChild(backdrop);
        window.setTimeout(() => search.focus(), 0);
    }

    private appendSchoolListTable(modal: HTMLElement, rows: SchoolSummaryRow[]): void {
        const scroller = document.createElement("div");
        scroller.className = "detail-table-scroller school-list-scroller";

        const table = document.createElement("table");
        table.className = "detail-table school-list-table";
        const thead = document.createElement("thead");
        const headerRow = document.createElement("tr");
        ["Codigo Local", "Nombre Colegio", "Riesgo", "Region", "Provincia", "Distrito", "Accion"].forEach((label: string) => {
            const th = document.createElement("th");
            th.textContent = label;
            headerRow.appendChild(th);
        });
        thead.appendChild(headerRow);
        table.appendChild(thead);

        const tbody = document.createElement("tbody");
        if (!rows.length) {
            const row = document.createElement("tr");
            const cell = document.createElement("td");
            cell.colSpan = 7;
            cell.textContent = "No hay colegios para el contexto seleccionado.";
            row.appendChild(cell);
            tbody.appendChild(row);
        } else {
            rows.forEach((school: SchoolSummaryRow) => {
                const row = document.createElement("tr");
                [
                    school.codigoLocal,
                    school.nombreLocal,
                    school.nivelRiesgo,
                    school.region,
                    school.provincia,
                    school.distrito
                ].forEach((value: string) => {
                    const td = document.createElement("td");
                    td.textContent = value || "-";
                    row.appendChild(td);
                });

                const actionCell = document.createElement("td");
                const action = document.createElement("button");
                action.className = "school-action-button";
                action.type = "button";
                action.textContent = "Ver Intervenciones";
                action.onclick = () => this.openSchoolInterventionsModal(school);
                actionCell.appendChild(action);
                row.appendChild(actionCell);
                tbody.appendChild(row);
            });
        }

        table.appendChild(tbody);
        scroller.appendChild(table);
        modal.appendChild(scroller);
    }

    private appendSchoolListPagination(modal: HTMLElement, totalPages: number): void {
        const pagination = document.createElement("div");
        pagination.className = "detail-pagination";

        const previous = document.createElement("button");
        previous.type = "button";
        previous.textContent = "Anterior";
        previous.disabled = this.detailPage <= 1;
        previous.onclick = () => {
            this.detailPage = Math.max(1, this.detailPage - 1);
            this.refreshSchoolModals();
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
            this.refreshSchoolModals();
        };
        pagination.appendChild(next);

        modal.appendChild(pagination);
    }

    private appendSchoolInterventionsModal(schoolKey: string): void {
        const school = this.getSchoolSummaryRows().find((row: SchoolSummaryRow) => this.normalizeText(row.codigoLocal) === schoolKey);
        if (!school) {
            this.selectedSchoolKey = null;
            return;
        }

        const backdrop = document.createElement("div");
        backdrop.className = "detail-modal-backdrop school-interventions-modal-backdrop";
        backdrop.onclick = (event: MouseEvent) => {
            if (event.target === backdrop) {
                this.closeSchoolInterventionsModal();
            }
        };

        const modal = document.createElement("section");
        modal.className = "detail-modal school-interventions-modal";

        const header = document.createElement("header");
        header.className = "school-interventions-header";
        const titleGroup = document.createElement("div");
        const eyebrow = document.createElement("span");
        eyebrow.textContent = school.codigoLocal;
        titleGroup.appendChild(eyebrow);
        const title = document.createElement("h2");
        title.textContent = school.nombreLocal;
        titleGroup.appendChild(title);

        const meta = document.createElement("div");
        meta.className = "school-interventions-meta";
        [
            ["Riesgo", school.nivelRiesgo],
            ["Region", school.region],
            ["Provincia", school.provincia],
            ["Distrito", school.distrito]
        ].forEach(([label, value]: string[]) => {
            const item = document.createElement("div");
            const itemLabel = document.createElement("small");
            itemLabel.textContent = label;
            item.appendChild(itemLabel);
            const itemValue = document.createElement("strong");
            itemValue.textContent = value || "-";
            item.appendChild(itemValue);
            meta.appendChild(item);
        });
        titleGroup.appendChild(meta);
        header.appendChild(titleGroup);

        const closeIcon = document.createElement("button");
        closeIcon.className = "detail-modal-close";
        closeIcon.type = "button";
        closeIcon.textContent = "X";
        closeIcon.onclick = () => this.closeSchoolInterventionsModal();
        header.appendChild(closeIcon);
        modal.appendChild(header);

        const availableUnits = this.getSchoolUnits(schoolKey);
        if (!availableUnits.includes(this.selectedSchoolTab)) {
            this.selectedSchoolTab = availableUnits[0] || "UGRD";
        }

        const tabs = document.createElement("div");
        tabs.className = "school-unit-tabs";
        (["UGRD", "UGSC", "UGEO", "UGM", "UGME"] as UnitKpiName[]).forEach((unit: UnitKpiName) => {
            const button = document.createElement("button");
            const hasData = availableUnits.includes(unit);
            button.type = "button";
            button.className = `school-unit-tab${unit === this.selectedSchoolTab ? " school-unit-tab-active" : ""}`;
            button.textContent = unit;
            button.disabled = !hasData;
            button.onclick = () => {
                this.selectedSchoolTab = unit;
                this.refreshSchoolModals();
            };
            tabs.appendChild(button);
        });
        modal.appendChild(tabs);

        const body = document.createElement("section");
        body.className = "school-interventions-body";
        const unitRows = this.getSchoolUnitRows(schoolKey, this.selectedSchoolTab);
        if (!unitRows.length) {
            const empty = document.createElement("p");
            empty.className = "school-interventions-empty";
            empty.textContent = "No hay intervenciones registradas para esta unidad gerencial.";
            body.appendChild(empty);
        } else if (this.selectedSchoolTab === "UGME") {
            this.appendSchoolUgmeInterventions(body, unitRows);
        } else {
            const summary = this.summarizeSchoolUnitRows(this.selectedSchoolTab, unitRows);
            const card = this.createUnitCard(summary, false);
            card.classList.add("school-intervention-unit-card");
            body.appendChild(card);
        }
        modal.appendChild(body);

        const footer = document.createElement("footer");
        footer.className = "detail-modal-footer";
        const backButton = document.createElement("button");
        backButton.type = "button";
        backButton.textContent = "Volver a colegios";
        backButton.onclick = () => this.closeSchoolInterventionsModal();
        footer.appendChild(backButton);
        modal.appendChild(footer);

        backdrop.appendChild(modal);
        this.target.appendChild(backdrop);
    }

    private appendSchoolUgmeInterventions(container: HTMLElement, rows: UnitDetailRow[]): void {
        const summary = this.summarizeSchoolUnitRows("UGME", rows);
        const hero = document.createElement("section");
        hero.className = "school-ugme-summary";
        this.getUnitKpiCells(summary).forEach((item: UnitKpiCell) => hero.appendChild(this.createKpiCell(item)));
        container.appendChild(hero);

        const sections = document.createElement("div");
        sections.className = "ugme-detail-sections school-ugme-sections";
        this.appendUgmeGroupSection(sections, "MODULOS PREFABRICADOS", this.getUgmeModularGroups(summary.grupos), "ugme-section-modulares");
        this.appendUgmeGroupSection(sections, "MOBILIARIO Y EQUIPAMIENTO", this.getUgmeMobiliarioGroups(summary.grupos), "ugme-section-mobiliario");
        container.appendChild(sections);
    }

    private getSchoolSummaryRows(): SchoolSummaryRow[] {
        const bySchool = new Map<string, SchoolSummaryRow>();

        this.streamingIndexes.detailRowsByUnit.forEach((rows: Map<string, UnitDetailRow>, unit: UnitKpiName) => {
            rows.forEach((row: UnitDetailRow) => {
                if (!this.detailRowMatchesInternalFilters(row)) {
                    return;
                }

                const key = this.normalizeText(row.codigoLocal);
                if (!key) {
                    return;
                }

                const existing = bySchool.get(key) || {
                    codigoLocal: row.codigoLocal || "-",
                    nombreLocal: row.nombreLocal || "-",
                    nivelRiesgo: row.nivelRiesgo || "-",
                    region: row.region || "-",
                    provincia: row.provincia || "-",
                    distrito: row.distrito || "-",
                    units: new Set<UnitKpiName>()
                };

                existing.units.add(unit);
                if ((!existing.nivelRiesgo || existing.nivelRiesgo === "-") && row.nivelRiesgo) {
                    existing.nivelRiesgo = row.nivelRiesgo;
                }
                bySchool.set(key, existing);
            });
        });

        return Array.from(bySchool.values())
            .sort((left: SchoolSummaryRow, right: SchoolSummaryRow) => {
                const byRegion = left.region.localeCompare(right.region);
                if (byRegion) {
                    return byRegion;
                }

                const byProvince = left.provincia.localeCompare(right.provincia);
                if (byProvince) {
                    return byProvince;
                }

                return left.codigoLocal.localeCompare(right.codigoLocal);
            });
    }

    private filterSchoolSummaryRows(rows: SchoolSummaryRow[]): SchoolSummaryRow[] {
        const searchText = this.normalizeText(this.detailSearchText);
        return rows.filter((row: SchoolSummaryRow) => {
            if (this.schoolListRegionFilter && !this.matchesRegion(row.region, this.schoolListRegionFilter)) {
                return false;
            }

            if (!searchText) {
                return true;
            }

            return [
                row.codigoLocal,
                row.nombreLocal,
                row.nivelRiesgo,
                row.region,
                row.provincia,
                row.distrito
            ].some((value: string) => this.normalizeText(value).includes(searchText));
        });
    }

    private getFirstSchoolUnit(row: SchoolSummaryRow): UnitKpiName | null {
        return (["UGRD", "UGSC", "UGEO", "UGM", "UGME"] as UnitKpiName[])
            .find((unit: UnitKpiName) => row.units.has(unit)) || null;
    }

    private getSchoolUnits(schoolKey: string): UnitKpiName[] {
        return (["UGRD", "UGSC", "UGEO", "UGM", "UGME"] as UnitKpiName[])
            .filter((unit: UnitKpiName) => this.getSchoolUnitRows(schoolKey, unit).length > 0);
    }

    private getSchoolUnitRows(schoolKey: string, unit: UnitKpiName): UnitDetailRow[] {
        return Array.from(this.streamingIndexes.detailRowsByUnit.get(unit)?.values() || [])
            .filter((row: UnitDetailRow) => this.normalizeText(row.codigoLocal) === schoolKey && this.detailRowMatchesInternalFilters(row))
            .sort((left: UnitDetailRow, right: UnitDetailRow) => (left.grupo || "").localeCompare(right.grupo || ""));
    }

    private summarizeSchoolUnitRows(unit: UnitKpiName, rows: UnitDetailRow[]): UnitKpiSummary {
        const solicitudes = new Set<string>();
        const solicitudesRevision = new Set<string>();
        const solicitudesCulminadas = new Set<string>();
        const regiones = new Set<string>();
        const colegios = new Set<string>();
        const grupos = new Map<string, GroupMetricSummary>();

        let montoIntervencion = 0;
        let montoAsignado = 0;
        let montoTransferencia = 0;
        let montoRetirado = 0;
        let beneficiarios = 0;
        let cantidad = 0;

        rows.forEach((row: UnitDetailRow) => {
            row.solicitudes.forEach((value: string) => solicitudes.add(value));
            row.solicitudesRevision.forEach((value: string) => solicitudesRevision.add(value));
            row.solicitudesCulminadas.forEach((value: string) => solicitudesCulminadas.add(value));
            regiones.add(this.normalizeText(row.region));
            colegios.add(this.normalizeText(row.codigoLocal));
            montoIntervencion += row.montoIntervencion;
            montoAsignado += row.montoAsignado;
            montoTransferencia += row.montoTransferencia;
            montoRetirado += row.montoRetirado;
            beneficiarios += row.beneficiarios;
            cantidad += row.cantidad;

            if (unit === "UGME" && row.grupo) {
                const existing = grupos.get(row.grupo) || {
                    grupo: row.grupo,
                    montoIntervencion: 0,
                    beneficiarios: 0,
                    cantidad: 0,
                    colegios: 0,
                    regiones: 0,
                    filas: 0
                };
                existing.montoIntervencion += row.montoIntervencion;
                existing.beneficiarios += row.beneficiarios;
                existing.cantidad += row.cantidad;
                existing.colegios = 1;
                existing.regiones = 1;
                existing.filas += 1;
                grupos.set(row.grupo, existing);
            }
        });

        return {
            unit,
            unitKey: unit,
            unidad: unit,
            label: unit,
            filas: rows.length,
            solicitudes: solicitudes.size,
            regiones: regiones.size,
            colegios: colegios.size,
            montoIntervencion,
            montoAsignado,
            montoTransferencia,
            montoRetirado,
            beneficiarios,
            solicitudesRevision: solicitudesRevision.size,
            solicitudesCulminadas: solicitudesCulminadas.size,
            colegiosConTransferencia: montoTransferencia > 0 ? 1 : 0,
            colegiosConRetiro: montoRetirado > 0 ? 1 : 0,
            grupos: unit === "UGME"
                ? ugmeGrupoOrder.map((grupo: string) => grupos.get(grupo) || {
                    grupo,
                    montoIntervencion: 0,
                    beneficiarios: 0,
                    cantidad: 0,
                    colegios: 0,
                    regiones: 0,
                    filas: 0
                })
                : [],
            registros: rows.length
        };
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
            this.refreshDetailModal();
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
        window.setTimeout(() => {
            const activeSearch = backdrop.querySelector(".detail-search-input") as HTMLInputElement | null;
            if (activeSearch) {
                activeSearch.focus();
                activeSearch.setSelectionRange(activeSearch.value.length, activeSearch.value.length);
            }
        }, 0);
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
            this.refreshDetailModal();
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
            this.refreshDetailModal();
        };
        pagination.appendChild(next);

        modal.appendChild(pagination);
    }

    private getUnitSchoolDetails(unit: UnitKpiName): UnitDetailRow[] {
        const cacheKey = `${this.lastDatasetSignature}|${unit}`;
        const cachedRows = this.detailCacheByUnit.get(cacheKey);
        if (cachedRows) {
            return cachedRows;
        }

        const rows = Array.from(this.streamingIndexes.detailRowsByUnit.get(unit)?.values() || [])
            .sort((left: UnitDetailRow, right: UnitDetailRow) => {
                const byRegion = left.region.localeCompare(right.region);
                if (byRegion) {
                    return byRegion;
                }

                return left.codigoLocal.localeCompare(right.codigoLocal);
            });

        this.cacheDetailRows(cacheKey, rows);
        return rows;
    }

    private filterDetailRows(rows: UnitDetailRow[]): UnitDetailRow[] {
        const searchText = this.normalizeText(this.detailSearchText);
        const internalFilteredRows = this.hasDashboardFilters()
            ? rows.filter((row: UnitDetailRow) => this.detailRowMatchesInternalFilters(row))
            : rows;
        const regionRows = this.selectedRegion
            ? internalFilteredRows.filter((row: UnitDetailRow) => this.isRegionNameSelected(row.region))
            : internalFilteredRows;

        if (!searchText) {
            return regionRows;
        }

        return regionRows.filter((row: UnitDetailRow) => [
            row.codigoLocal,
            row.nombreLocal,
            row.region,
            row.provincia,
            row.distrito,
            row.grupo || ""
        ].some((value: string) => this.normalizeText(value).includes(searchText)));
    }

    private detailRowMatchesInternalFilters(row: UnitDetailRow): boolean {
        const filters = this.internalFilters;

        if (!this.detailRowMatchesSelectedRisk(row)) {
            return false;
        }

        if (filters.region && this.normalizeText(row.region) !== this.normalizeText(filters.region)) {
            return false;
        }

        if (filters.provincia && this.normalizeText(row.provincia) !== this.normalizeText(filters.provincia)) {
            return false;
        }

        if (filters.distrito && this.normalizeText(row.distrito) !== this.normalizeText(filters.distrito)) {
            return false;
        }

        if (filters.codigoLocal && this.normalizeText(row.codigoLocal) !== this.normalizeText(filters.codigoLocal)) {
            return false;
        }

        if (filters.nombreLocal && this.normalizeText(row.nombreLocal) !== this.normalizeText(filters.nombreLocal)) {
            return false;
        }

        return true;
    }

    private recordMatchesSelectedRisk(record: SourceRecord): boolean {
        if (!this.hasRiskFilter()) {
            return true;
        }

        return this.normalizeRiskLevelForFilter(record.nivelRiesgo) === this.selectedRiskView;
    }

    private detailRowMatchesSelectedRisk(row: UnitDetailRow): boolean {
        if (!this.hasRiskFilter()) {
            return true;
        }

        return this.normalizeRiskLevelForFilter(row.nivelRiesgo) === this.selectedRiskView;
    }

    private normalizeRiskLevelForFilter(value: PrimitiveValue): RiskLabel {
        const normalized = this.normalizeText(value)
            .replace(/\s+/g, " ")
            .trim();

        if (normalized === "MUY ALTO" || normalized === "MUYALTO" || (normalized.includes("MUY") && normalized.includes("ALTO"))) {
            return "Muy Alto";
        }

        if (normalized === "ALTO" || normalized.includes("ALTO")) {
            return "Alto";
        }

        if (normalized === "MEDIO" || normalized.includes("MEDIO")) {
            return "Medio";
        }

        if (normalized === "BAJO" || normalized.includes("BAJO")) {
            return "Bajo";
        }

        return "Sin Clasificación";
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
                { label: "Monto", value: (row: UnitDetailRow) => this.formatMillions(row.montoIntervencion) }
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

    private appendFinalLoadingPanel(diagnostics: TableDiagnostics): void {
        const panel = document.createElement("section");
        panel.className = "unit-summary-view";

        const title = document.createElement("h1");
        title.textContent = "RESUMEN POR UNIDADES GERENCIALES";
        panel.appendChild(title);

        const message = document.createElement("p");
        message.className = "loading-message";
        message.textContent = diagnostics.accumulatedFilas
            ? "Preparando vista..."
            : "Cargando informacion...";
        panel.appendChild(message);

        this.target.appendChild(panel);
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
                ["Monto", this.formatMillions(kpi.montoIntervencion)]
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
            "Monto"
        ]);
        const tbody = document.createElement("tbody");

        grupos.forEach((grupo: GroupMetricSummary) => {
            this.appendTableRow(tbody, [
                grupo.grupo,
                this.formatInteger(grupo.cantidad),
                this.formatInteger(grupo.beneficiarios),
                this.formatMillions(grupo.montoIntervencion)
            ]);
        });

        table.appendChild(tbody);
        card.appendChild(table);
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

    private getCachedAnalytics(datasetSignature: string): AnalyticsEngine | null {
        const cached = this.analyticsCacheBySignature.get(datasetSignature);
        if (!cached) {
            return null;
        }

        cached.lastUsed = ++this.cacheUseSequence;
        return cached.analytics;
    }

    private cacheAnalytics(datasetSignature: string, analytics: AnalyticsEngine, rowCount: number): void {
        this.analyticsCacheBySignature.set(datasetSignature, {
            analytics,
            datasetSignature,
            rowCount,
            lastUsed: ++this.cacheUseSequence
        });
        this.pruneAnalyticsCache();
    }

    private pruneAnalyticsCache(): void {
        if (this.analyticsCacheBySignature.size <= this.maxAnalyticsCacheEntries) {
            return;
        }

        const entries = Array.from(this.analyticsCacheBySignature.values())
            .sort((left: AnalyticsCacheEntry, right: AnalyticsCacheEntry) => left.lastUsed - right.lastUsed);
        const entriesToDelete = entries.slice(0, this.analyticsCacheBySignature.size - this.maxAnalyticsCacheEntries);

        entriesToDelete.forEach((entry: AnalyticsCacheEntry) => {
            this.analyticsCacheBySignature.delete(entry.datasetSignature);
        });
    }

    private cacheDetailRows(cacheKey: string, rows: UnitDetailRow[]): void {
        this.detailCacheByUnit.set(cacheKey, rows);

        if (this.detailCacheByUnit.size <= this.maxDetailCacheEntries) {
            return;
        }

        const keysToDelete = Array.from(this.detailCacheByUnit.keys())
            .slice(0, this.detailCacheByUnit.size - this.maxDetailCacheEntries);
        keysToDelete.forEach((key: string) => this.detailCacheByUnit.delete(key));
    }

    private buildDatasetSignature(columns: DataViewMetadataColumn[], hasMoreData: boolean): string {
        if (!this.enableAnalyticsCache) {
            return [
                `rows=${this.accumulatedRowCount}`,
                `segment=${hasMoreData}`,
                this.buildColumnSignature(columns)
            ].join("|");
        }

        return [
            `rows=${this.accumulatedRowCount}`,
            `segment=${hasMoreData}`,
            this.buildColumnSignature(columns),
            this.buildRecordFingerprint(this.accumulatedRows)
        ].join("|");
    }

    private buildIncomingDataSignature(records: SourceRecord[], columns: DataViewMetadataColumn[], hasMoreData: boolean): string {
        return [
            `rows=${records.length}`,
            `segment=${hasMoreData}`,
            this.buildColumnSignature(columns),
            this.buildRecordBoundarySignature(records)
        ].join("|");
    }

    private buildRecordBoundarySignature(records: SourceRecord[]): string {
        if (!records.length) {
            return "empty";
        }

        return [
            `first=${this.recordKey(records[0])}`,
            `last=${this.recordKey(records[records.length - 1])}`
        ].join("|");
    }

    private buildColumnSignature(columns: DataViewMetadataColumn[]): string {
        return columns
            .map((column: DataViewMetadataColumn) => column.queryName || column.displayName || "")
            .join(",");
    }

    private buildRecordFingerprint(records: SourceRecord[]): string {
        if (!records.length) {
            return "empty";
        }

        let xorHash = 0;
        let sumHash = 0;
        let amountTotal = 0;
        let beneficiaryTotal = 0;

        for (const record of records) {
            const rowHash = this.hashString([
                record.region,
                record.provincia,
                record.distrito,
                record.codigoLocal,
                record.idSolicitud,
                record.unidadGerencial,
                record.estadoSolicitud,
                record.nivelRiesgo,
                record.montoInversion
            ].join("|"));
            xorHash ^= rowHash;
            sumHash = (sumHash + (rowHash >>> 0)) >>> 0;
            amountTotal += record.montoInversion || 0;
            beneficiaryTotal += record.beneficiarios || 0;
        }

        return [
            `xor=${xorHash >>> 0}`,
            `sum=${sumHash}`,
            `amount=${Math.round(amountTotal * 100) / 100}`,
            `benef=${beneficiaryTotal}`
        ].join("|");
    }

    private hashString(value: string): number {
        let hash = 2166136261;

        for (let index = 0; index < value.length; index += 1) {
            hash ^= value.charCodeAt(index);
            hash = Math.imul(hash, 16777619);
        }

        return hash;
    }

    private resetAccumulationIfDatasetChanged(
        records: SourceRecord[],
        columns: DataViewMetadataColumn[],
        hasMoreData: boolean,
        operationKind?: powerbi.VisualDataChangeOperationKind
    ): boolean {
        const columnSignature = this.buildColumnSignature(columns);
        const incomingSignature = this.buildIncomingDataSignature(records, columns, hasMoreData);
        const isCreateOperation = operationKind === powerbi.VisualDataChangeOperationKind.Create;
        const columnsChanged = !!this.lastColumnSignature && columnSignature !== this.lastColumnSignature;
        const startsNewSegmentedLoad = this.allDataLoaded && hasMoreData && !this.isFetching;
        const incomingChangedAfterComplete = this.allDataLoaded &&
            !this.isFetching &&
            !!this.lastIncomingDataSignature &&
            incomingSignature !== this.lastIncomingDataSignature;

        if (columnsChanged) {
            this.analyticsCacheBySignature = new Map<string, AnalyticsCacheEntry>();
        }

        if (isCreateOperation || columnsChanged || startsNewSegmentedLoad || incomingChangedAfterComplete) {
            this.resetStreamingState();
            this.fetchCount = 0;
            this.analyticsCache = null;
            this.incrementalAnalytics = null;
            this.selectedRegionAnalyticsCache = null;
            this.selectedRegionAnalyticsCacheKey = "";
            this.lastDatasetSignature = "";
            this.lastAccumulatedIncomingSignature = "";
            this.segmentLoadStartTime = undefined;
            this.allDataLoaded = false;
            this.selectedRegion = null;
            this.autoFocusedRegion = null;
            this.suppressedAutoFocusSignature = "";
            this.activeMapDrillKey = "";
        }

        this.lastColumnSignature = columnSignature;
        this.lastIncomingDataSignature = incomingSignature;

        return isCreateOperation || columnsChanged || startsNewSegmentedLoad || incomingChangedAfterComplete;
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

    private internDimensionString(value: string): string {
        const cached = this.dimensionStringPool.get(value);
        if (cached !== undefined) {
            return cached;
        }

        this.dimensionStringPool.set(value, value);
        return value;
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
        const amount = Number.isFinite(value) ? value : 0;

        if (amount >= 1_000_000) {
            const millions = Math.floor(amount / 1_000_000);
            const formattedMillions = millions.toLocaleString("en-US", {
                maximumFractionDigits: 0,
                minimumFractionDigits: 0
            });
            return `S/ ${formattedMillions} M`;
        }

        return `S/ ${amount.toLocaleString("en-US", {
            maximumFractionDigits: 0,
            minimumFractionDigits: 0
        })}`;
    }

    private formatMillions(value: number): string {
        return this.formatCurrencyNoDecimals(value);
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

    private formatPercentFromRatio(value: number): string {
        if (!Number.isFinite(value)) {
            return "0.00%";
        }

        return `${(value * 100).toFixed(2)}%`;
    }

    private formatRoundedPercentFromRatio(value: number): string {
        if (!Number.isFinite(value)) {
            return "0%";
        }

        return `${Math.round(value * 100)}%`;
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
            this.valueKey(record.codigoLocal),
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
            record.montoRetirado.toString()
        ].join("|");
    }
}
