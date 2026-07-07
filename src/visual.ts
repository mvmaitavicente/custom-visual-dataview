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
type RiskMapView = "Total" | RiskLevel;

const riskMapViews: RiskMapView[] = ["Total", "Bajo", "Medio", "Alto", "Muy Alto"];

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

function getRiskHeatColor(ratio: number): string {
    if (!Number.isFinite(ratio) || ratio <= 0) {
        return "#e5e7eb";
    }

    const clamped = Math.max(0, Math.min(1, ratio));

    if (clamped <= 0.25) {
        return "#22c55e";
    }

    if (clamped <= 0.5) {
        return "#facc15";
    }

    if (clamped <= 0.75) {
        return "#fb923c";
    }

    return "#ef4444";
}

function getRiskGradientColor(score: number): string {
    if (!Number.isFinite(score) || score < 0) {
        return "#e5e7eb";
    }

    const stops = [
        { at: 0, color: [187, 247, 208] },
        { at: 0.32, color: [190, 242, 100] },
        { at: 0.55, color: [253, 224, 71] },
        { at: 0.78, color: [251, 146, 60] },
        { at: 1, color: [239, 68, 68] }
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
    private detailSearchText = "";
    private detailPage = 1;
    private detailPageSize = 20;
    private detailCacheByUnit: Map<string, UnitDetailRow[]> = new Map<string, UnitDetailRow[]>();
    private readonly maxAnalyticsCacheEntries = 4;
    private readonly maxDetailCacheEntries = 1;
    private readonly enableAnalyticsCache = false;
    private readonly enableInterimLoadingRender = false;
    private readonly enableSelectedRegionCache = false;
    private readonly enableSegmentAutoFetch = true;
    private readonly useCurrentDataViewOnly = false;
    private readonly enableRuntimeLogs = false;
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
        this.removeLeafletMap();
        window.removeEventListener("keydown", this.handleKeydown);
    }

    public update(options: VisualUpdateOptions): void {
        this.updateCount += 1;
        const updateStart = performance.now();
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
            this.updatePerformancePanel(diagnostics);
            this.updatePerformanceMini(diagnostics);
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

        const diagnostics = this.buildDiagnostics(dataViews, fetchWaitMs, options.operationKind);
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
        operationKind?: powerbi.VisualDataChangeOperationKind
    ): TableDiagnostics {
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
        const datasetChanged = this.resetAccumulationIfDatasetChanged(records, columns, hasMoreData, operationKind);

        const incomingDataSignature = this.lastIncomingDataSignature;
        const accumulateResult = measure("accumulateRows", () => this.accumulateRecords(records, hasMoreData, incomingDataSignature));
        performanceMetrics.accumulateRowsMs = accumulateResult.ms;
        this.accumulatedRowCount = this.accumulatedRowCount || records.length;
        this.logDataFlow("after-accumulate", {
            hasMoreData,
            datasetChanged,
            segmentRows: records.length,
            accumulatedRows: this.accumulatedRowCount,
            isFetching: this.isFetching,
            allDataLoaded: this.allDataLoaded,
            operationKind,
            incomingSignature: incomingDataSignature,
            lastDatasetSignature: this.lastDatasetSignature
        });
        if (this.isFetching && fetchWaitMs !== undefined) {
            this.isFetching = false;
        }

        if (hasMoreData && this.enableSegmentAutoFetch) {
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
            performanceMetrics.segmentLoadMs = this.segmentLoadStartTime === undefined
                ? 0
                : performance.now() - this.segmentLoadStartTime;

            const interimAnalyticsResult = this.enableInterimLoadingRender && datasetChanged && this.accumulatedRowCount
                ? measure("buildInterimAnalyticsEngine", () => this.incrementalAnalytics || AnalyticsEngine.build(this.accumulatedRows))
                : null;
            const interimAnalytics = interimAnalyticsResult?.value || null;
            if (interimAnalyticsResult) {
                performanceMetrics.buildAnalyticsEngineMs = interimAnalyticsResult.ms;
                performanceMetrics.analyticsRebuilt = true;

                const buildMetrics = interimAnalyticsResult.value.getBuildMetrics();
                performanceMetrics.buildFlatIndexesMs = buildMetrics.buildFlatIndexesMs;
                performanceMetrics.buildRelationIndexesMs = buildMetrics.buildRelationIndexesMs;
                performanceMetrics.lazySampleQueryMs = buildMetrics.lazySampleQueryMs;
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
        this.logDataFlow("final-render", {
            datasetSignature,
            accumulatedRows: this.accumulatedRowCount,
            analyticsCacheHit,
            analyticsRebuilt,
            buildAnalyticsEngineMs: performanceMetrics.buildAnalyticsEngineMs
        });

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
            this.appendPerformanceMini(diagnostics);
            return;
        }

        if (this.hasMissingRole(diagnostics.columnIndexes)) {
            this.appendMessage("Missing one or more required roles for unit KPI cards.");
            this.appendPerformanceMini(diagnostics);
            return;
        }

        if (diagnostics.isLoading && diagnostics.analytics) {
            this.reconcileInternalFilters();
            const renderAnalytics = this.getInternalFilteredAnalytics(diagnostics.analytics);
            if (this.currentView === "ugme") {
                this.appendUgmeView(renderAnalytics);
            } else {
                this.appendSummaryView(renderAnalytics);
            }

            this.appendLoadingStatusBadge(diagnostics);
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

        this.reconcileInternalFilters();
        const renderAnalytics = this.getInternalFilteredAnalytics(diagnostics.analytics);
        if (this.currentView === "ugme") {
            this.appendUgmeView(renderAnalytics);
        } else {
            this.appendSummaryView(renderAnalytics);
        }

        if (this.isDetailModalOpen && this.detailModalUnit) {
            this.appendDetailModal(this.detailModalUnit);
        }

        this.appendPerformanceMini(diagnostics);
    }

    private appendSummaryView(engine: AnalyticsEngine): void {
        const view = document.createElement("section");
        view.className = "dashboard-root with-internal-filters";
        this.appendInternalFilterBar(view);

        const layout = document.createElement("div");
        layout.className = "main-dashboard-layout";

        const mapPanel = document.createElement("section");
        mapPanel.className = "map-panel";
        this.appendPeruRiskMap(mapPanel, engine);
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

        const title = document.createElement("h1");
        title.textContent = this.selectedRegion
            ? `RESUMEN POR UNIDADES GERENCIALES - ${this.selectedRegion}`
            : "RESUMEN POR UNIDADES GERENCIALES";
        unitsPanel.appendChild(title);

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

    private getInternalFilterOptions(key: InternalFilterKey): string[] {
        const values = new Set<string>();

        this.accumulatedRows.forEach((record: SourceRecord) => {
            if (!this.recordMatchesInternalFilters(record, key)) {
                return;
            }

            const value = record[key];
            if (value) {
                values.add(value);
            }
        });

        return Array.from(values).sort((left: string, right: string) => left.localeCompare(right));
    }

    private recordMatchesInternalFilters(record: SourceRecord, ignoreKey?: InternalFilterKey): boolean {
        const filters = this.internalFilters;

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
                this.recordMatchesInternalFilters(record, key)
            ));

            if (!exists) {
                this.internalFilters[key] = "";
            }
        });
    }

    private getInternalFilteredAnalytics(baseEngine: AnalyticsEngine): AnalyticsEngine {
        if (!this.hasInternalFilters() || !this.accumulatedRows.length) {
            return baseEngine;
        }

        const filterKey = this.getInternalFilterCacheKey();
        const cached = this.internalFilterAnalyticsCache.get(filterKey);
        if (cached) {
            return cached;
        }

        const filteredRows = this.accumulatedRows.filter((record: SourceRecord) => this.recordMatchesInternalFilters(record));
        const filteredEngine = AnalyticsEngine.build(filteredRows);
        this.internalFilterAnalyticsCache.set(filterKey, filteredEngine);

        if (this.internalFilterAnalyticsCache.size > 4) {
            const oldestKey = this.internalFilterAnalyticsCache.keys().next().value;
            if (oldestKey) {
                this.internalFilterAnalyticsCache.delete(oldestKey);
            }
        }

        return filteredEngine;
    }

    private getInternalFilterCacheKey(): string {
        return [
            this.lastDatasetSignature,
            this.normalizeText(this.internalFilters.region),
            this.normalizeText(this.internalFilters.provincia),
            this.normalizeText(this.internalFilters.distrito),
            this.normalizeText(this.internalFilters.codigoLocal),
            this.normalizeText(this.internalFilters.nombreLocal)
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
                    direction: "top",
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
        map.fitBounds(focusBounds, { padding: [20, 20] });
        map.invalidateSize();
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
            { label: "Ver Provincias", mode: "provincias" as MapDrillMode },
            { label: "Ver Distritos", mode: "distritos" as MapDrillMode },
            { label: "Ver Colegios", mode: "colegios" as MapDrillMode }
        ].forEach((option: { label: string; mode: MapDrillMode }) => {
            const button = document.createElement("button");
            button.type = "button";
            button.textContent = option.label;
            button.onclick = () => {
                this.hideMapContextMenu();
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
            fillOpacity: selected ? 0.82 : 0.68,
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
            const ratio = mapHeatScale.maxTotalColegios
                ? aggregate.totalColegios / mapHeatScale.maxTotalColegios
                : 0;
            return getRiskHeatColor(ratio);
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

        this.getRegionTooltipRows(aggregate).forEach(([label, value, className]: string[]) => {
            this.appendTooltipRow(tooltip, label, value, className);
        });

        return tooltip;

        [
            ["N° Colegios", this.formatInteger(aggregate?.totalColegios || 0)],
            ["N° Provincias", this.formatInteger(aggregate?.totalProvincias || 0)],
            ["N° Colegio - Riesgo Bajo", this.formatInteger(aggregate?.riskBreakdown.bajo || 0)],
            ["N° Colegio - Riesgo Medio", this.formatInteger(aggregate?.riskBreakdown.medio || 0)],
            ["N° Colegio - Riesgo Alto", this.formatInteger(aggregate?.riskBreakdown.alto || 0)],
            ["N° Colegio - Riesgo Muy Alto", this.formatInteger(aggregate?.riskBreakdown.muyAlto || 0)]
        ].forEach(([label, value]: string[]) => {
            this.appendTooltipRow(tooltip, label, value);
        });

        return tooltip;
    }

    private getRegionTooltipRows(aggregate: RegionAggregate | undefined): string[][] {
        const totalColegios = aggregate?.totalColegios || 0;

        if (this.selectedRiskView !== "Total") {
            const selectedRiskCount = aggregate
                ? this.getRiskValue(aggregate.riskBreakdown, this.selectedRiskView)
                : 0;

            return [
                ["Colegios", this.formatInteger(totalColegios), "region-tooltip-row-main"],
                [this.selectedRiskView, this.formatInteger(selectedRiskCount), this.getRiskTooltipClass(this.selectedRiskView)],
                ["Porcentaje", this.formatPercentFromRatio(this.getSelectedRiskPercentage(aggregate)), "region-tooltip-percent"]
            ];
        }

        return [
            ["Colegios", this.formatInteger(totalColegios), "region-tooltip-row-main"],
            ["Provincias", this.formatInteger(aggregate?.totalProvincias || 0), "region-tooltip-row-main"],
            ["Bajo", this.formatInteger(aggregate?.riskBreakdown.bajo || 0), "risk-low"],
            ["Medio", this.formatInteger(aggregate?.riskBreakdown.medio || 0), "risk-medium"],
            ["Alto", this.formatInteger(aggregate?.riskBreakdown.alto || 0), "risk-high"],
            ["Muy alto", this.formatInteger(aggregate?.riskBreakdown.muyAlto || 0), "risk-critical"]
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
        hint.textContent = this.selectedRiskView === "Total"
            ? "Selecciona un riesgo"
            : `Top por % ${this.selectedRiskView}`;
        panel.appendChild(hint);

        const list = document.createElement("ol");
        list.className = "map-ranking-list";
        this.getMapRankingRows(engine).forEach((row: { region: string; percentage: number; score: number }) => {
            const item = document.createElement("li");

            const swatch = document.createElement("span");
            swatch.style.background = getRiskGradientColor(row.score);
            item.appendChild(swatch);

            const label = document.createElement("b");
            label.textContent = this.shortRegionLabel(row.region);
            item.appendChild(label);

            const value = document.createElement("em");
            value.textContent = this.formatPercentFromRatio(row.percentage);
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

        const cacheKey = `${this.lastDatasetSignature}|${this.normalizeRegionForMap(this.selectedRegion)}`;
        if (this.enableSelectedRegionCache && this.selectedRegionAnalyticsCache && this.selectedRegionAnalyticsCacheKey === cacheKey) {
            return this.selectedRegionAnalyticsCache;
        }

        const selectedKey = this.normalizeRegionForMap(this.selectedRegion);
        const filteredEngine = this.streamingIndexes.regionAnalytics.get(selectedKey) || AnalyticsEngine.createEmpty();
        filteredEngine.finalizeBuild();
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

    private appendUgmeView(engine: AnalyticsEngine): void {
        const view = document.createElement("section");
        view.className = "dashboard-root ugme-dashboard-view with-internal-filters";
        this.appendInternalFilterBar(view);

        const layout = document.createElement("div");
        layout.className = "main-dashboard-layout";

        const mapPanel = document.createElement("section");
        mapPanel.className = "map-panel";
        this.appendPeruRiskMap(mapPanel, engine);
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
            { icon: "RG", label: "N° Regiones", value: this.formatInteger(kpi.regiones) },
            { icon: "CL", label: "N° Colegios", value: this.formatInteger(kpi.colegios) },
            { icon: "S/", label: "Monto", value: this.formatMillions(kpi.montoIntervencion) }
        ].forEach((item: UnitKpiCell) => kpiGrid.appendChild(this.createUgmeHeroKpi(item)));
        hero.appendChild(kpiGrid);
        card.appendChild(hero);

        this.appendUgmeGroupSection(card, "MODULARES", this.getUgmeModularGroups(kpi.grupos), "ugme-section-modulares");
        this.appendUgmeGroupSection(card, "MOBILIARIO Y EQUIPAMIENTO", this.getUgmeMobiliarioGroups(kpi.grupos), "ugme-section-mobiliario");

        return card;
    }

    private createUgmeHeroKpi(item: UnitKpiCell): HTMLElement {
        const card = document.createElement("article");
        card.className = "ugme-hero-kpi";

        const icon = document.createElement("span");
        icon.textContent = item.icon;
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
                { icon: "MI", label: "Monto", value: this.formatNumber(kpi.montoIntervencion) }
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
                { icon: "MI", label: "Monto", value: this.formatMillions(kpi.montoIntervencion) }
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

        const grid = document.createElement("div");
        grid.className = "ugme-detail-group-grid";
        grupos.forEach((grupo: GroupMetricSummary) => grid.appendChild(this.createUgmeDetailGroupCard(grupo)));
        section.appendChild(grid);
        container.appendChild(section);
    }

    private createUgmeDetailGroupCard(grupo: GroupMetricSummary): HTMLElement {
        const card = document.createElement("article");
        const groupClass = this.getUgmeGroupClass(grupo.grupo);
        card.className = `ugme-detail-group-card ${groupClass}`;

        const header = document.createElement("div");
        header.className = "ugme-detail-group-header";

        const icon = document.createElement("span");
        icon.textContent = this.getUgmeGroupIcon(grupo.grupo);
        header.appendChild(icon);

        const title = document.createElement("h4");
        title.textContent = grupo.grupo;
        header.appendChild(title);
        card.appendChild(header);

        [
            ["Cantidad", this.formatInteger(grupo.cantidad), "CA"],
            ["Beneficiarios", this.formatInteger(grupo.beneficiarios), "BE"],
            ["Monto", this.formatMillions(grupo.montoIntervencion), "S/"]
        ].forEach(([labelText, valueText, iconText]: string[]) => {
            const row = document.createElement("div");
            row.className = "ugme-detail-group-row";

            const label = document.createElement("span");
            label.textContent = iconText;
            row.appendChild(label);

            const text = document.createElement("small");
            text.textContent = labelText;
            row.appendChild(text);

            const value = document.createElement("strong");
            value.textContent = valueText;
            row.appendChild(value);
            card.appendChild(row);
        });

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
            "ESCUELA MODULAR": "ED",
            "AULA MODULAR": "AU",
            "KIT DE PARARRAYO": "KP",
            "MODULO SS.HH.": "SS",
            "REDES COMPLEMENTARIAS": "RC",
            MOBILIARIO: "MO",
            EQUIPAMIENTO: "EQ",
            LABORATORIOS: "LA",
            TALLERES: "TA"
        };

        return icons[grupo] || "ME";
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
        const regionRows = this.selectedRegion
            ? rows.filter((row: UnitDetailRow) => this.isRegionNameSelected(row.region))
            : rows;

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

    private appendLoadingStatusBadge(diagnostics: TableDiagnostics): void {
        const badge = document.createElement("div");
        badge.className = "loading-status-badge";
        badge.textContent = `Actualizando datos: ${this.formatInteger(diagnostics.accumulatedFilas)} filas`;
        this.target.appendChild(badge);
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
            ["% critico global", this.formatPercent(totals.porcentajeCritico)]
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
            "% critico"
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
                this.formatPercent(summary.porcentajeCritico)
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
            this.formatPercent(totals.porcentajeCritico)
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
            "% critico"
        ]);
        const tbody = document.createElement("tbody");

        engine.getUnidades().forEach((summary: BucketSummary) => {
            this.appendTableRow(tbody, [
                summary.label,
                summary.filas.toString(),
                summary.colegiosUnicos.toString(),
                summary.solicitudesUnicas.toString(),
                this.formatNumber(summary.montoTotal),
                this.formatPercent(summary.porcentajeCritico)
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

    private logDataFlow(label: string, details: Record<string, unknown>): void {
        if (!this.enableRuntimeLogs) {
            return;
        }

        // Runtime-only diagnostics for Power BI Desktop filtering/segment behavior.
        console.group(`DataView flow #${this.updateCount}: ${label}`);
        Object.entries(details).forEach(([key, value]: [string, unknown]) => {
            console.log(key, value);
        });
        console.groupEnd();
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

    private formatMillions(value: number): string {
        const millions = value / 1_000_000;
        const fractionDigits = Math.abs(millions) < 100 ? 1 : 0;

        return `S/ ${millions.toLocaleString(undefined, {
            maximumFractionDigits: fractionDigits,
            minimumFractionDigits: fractionDigits
        })} M`;
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
