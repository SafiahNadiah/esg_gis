require([
  'esri/Map',
  'esri/views/MapView',
  'esri/layers/FeatureLayer',
  'esri/Graphic',
  'esri/widgets/Legend',
  'esri/widgets/Expand',
  'esri/geometry/geometryEngine',
  'esri/layers/GraphicsLayer',
  'esri/rest/query',
  'esri/geometry/Polyline',
  'esri/geometry/Polygon',
  'esri/widgets/Sketch',
  'esri/geometry/projection'
], function (Map, MapView, FeatureLayer, Graphic, Legend, Expand, geometryEngine, GraphicsLayer, query, Polyline, Polygon, Sketch, projection) {

    // supabase client provided by index.html
    const supabase = window.supabaseClient;

    // Add external Air Quality FeatureServer layer
    const airQualityLayer = new FeatureLayer({
      url:
        'https://services5.arcgis.com/VoTFsdkrXCTvwVPC/arcgis/rest/services/Air_Quality/FeatureServer/0',
      title: 'Air Quality',
      outFields: ['*'],
      renderer: {
        type: 'class-breaks',
        field: 'AQI',
        classBreakInfos: [
          {
            minValue: 0,
            maxValue: 50,
            symbol: {
              type: 'simple-marker',
              size: 12,
              color: '#1a9641',
              outline: { color: 'white', width: 1 }
            },
            label: 'Good (0-50)'
          },
          {
            minValue: 51,
            maxValue: 100,
            symbol: {
              type: 'simple-marker',
              size: 12,
              color: '#ffffbf',
              outline: { color: 'white', width: 1 }
            },
            label: 'Moderate (51-100)'
          },
          {
            minValue: 101,
            maxValue: 200,
            symbol: {
              type: 'simple-marker',
              size: 12,
              color: '#fee090',
              outline: { color: 'white', width: 1 }
            },
            label: 'Poor (101-200)'
          },
          {
            minValue: 201,
            maxValue: 1000,
            symbol: {
              type: 'simple-marker',
              size: 12,
              color: '#d7191c',
              outline: { color: 'white', width: 1 }
            },
            label: 'Very Poor (200+)'
          }
        ]
      },
      popupTemplate: {
        title: 'Air Quality Monitoring',
        content: [
          {
            type: 'fields',
            fieldInfos: [
              { fieldName: 'AQI', label: 'Air Quality Index (AQI)' },
              { fieldName: 'PM2_5', label: 'PM2.5 (µg/m³)' },
              { fieldName: 'PM10', label: 'PM10 (µg/m³)' },
              { fieldName: 'NO2', label: 'NO2 (ppb)' },
              { fieldName: 'O3', label: 'O3 (ppb)' },
              { fieldName: 'timestamp', label: 'Timestamp' }
            ]
          }
        ]
      }
    });

    // Add external Industrial Emissions FeatureServer layer
    const industrialEmissionsLayer = new FeatureLayer({
      url:
        'https://services5.arcgis.com/VoTFsdkrXCTvwVPC/arcgis/rest/services/Industrial_Emissions/FeatureServer/0',
      title: 'Industrial Emissions',
      outFields: ['*'],
      renderer: {
        type: 'unique-value',
        field: 'compliance_status',
        defaultSymbol: {
          type: 'simple-marker',
          size: 10,
          color: '#cccccc',
          outline: { color: 'white', width: 1 }
        },
        uniqueValueInfos: [
          {
            value: 'Compliant',
            symbol: {
              type: 'simple-marker',
              size: 10,
              color: '#1a9641',
              outline: { color: 'white', width: 1 }
            },
            label: 'Compliant'
          },
          {
            value: 'Non-Compliant',
            symbol: {
              type: 'simple-marker',
              size: 10,
              color: '#d7191c',
              outline: { color: 'white', width: 1 }
            },
            label: 'Non-Compliant'
          }
        ],
        visualVariables: [
          {
            type: 'size',
            field: 'annual_CO2_ton',
            minDataValue: 0,
            maxDataValue: 100000,
            minSize: 8,
            maxSize: 25
          }
        ]
      },
      popupTemplate: {
        title: '{facility_name}',
        content: [
          {
            type: 'fields',
            fieldInfos: [
              { fieldName: 'facility_name', label: 'Facility Name' },
              { fieldName: 'compliance_status', label: 'Compliance Status' },
              { fieldName: 'annual_CO2_ton', label: 'Annual CO2 (tons)' },
              { fieldName: 'industry_sector', label: 'Industry Sector' },
              { fieldName: 'location', label: 'Location' },
              { fieldName: 'last_inspection_date', label: 'Last Inspection Date' },
              { fieldName: 'emissions_limit', label: 'Emissions Limit (tons)' }
            ]
          }
        ]
      }
    });

  const map = new Map({ basemap: 'gray-vector', layers: [airQualityLayer, industrialEmissionsLayer] });

  // Fire Hotspots layer: supports heatmap and point symbology
  const fireHotspotsPointRenderer = {
    type: 'simple',
    symbol: {
      type: 'simple-marker',
      color: 'red',
      outline: { color: 'rgba(255,255,255,0.6)', width: 0.5 }
    },
    visualVariables: [
      {
        type: 'size',
        field: 'brightness',
        minDataValue: 0,
        maxDataValue: 300,
        minSize: 6,
        maxSize: 36
      }
    ]
  };

  const fireHotspotsHeatmapRenderer = {
    type: 'heatmap',
    field: 'brightness',
    colorStops: [
      { ratio: 0, color: 'rgba(0,0,0,0)' },
      { ratio: 0.2, color: 'rgba(255,238,173,0.6)' },
      { ratio: 0.4, color: 'rgba(255,204,92,0.6)' },
      { ratio: 0.6, color: 'rgba(255,140,60,0.6)' },
      { ratio: 0.8, color: 'rgba(255,60,40,0.7)' },
      { ratio: 1, color: 'rgba(180,0,20,0.8)' }
    ],
    radius: 21.6,
    maxDensity: 0.06028163580246914,
    minDensity: 0
  };

  const fireHotspotsLayer = new FeatureLayer({
    url: 'https://services5.arcgis.com/VoTFsdkrXCTvwVPC/arcgis/rest/services/Fire_Hotspots/FeatureServer/0',
    title: 'Fire Hotspots',
    outFields: ['*'],
    visible: false,
    // default to heatmap; switch to point symbology by assigning the point renderer
    renderer: fireHotspotsHeatmapRenderer,
    popupTemplate: {
      title: 'Fire Hotspot',
      content: [
        {
          type: 'fields',
          fieldInfos: [
            { fieldName: 'brightness', label: 'Brightness' },
            { fieldName: 'confidence', label: 'Confidence' },
            { fieldName: 'frp', label: 'Fire Radiative Power (FRP)' },
            { fieldName: 'acq_date', label: 'Acquisition Date' },
            { fieldName: 'acq_time', label: 'Acquisition Time' }
          ]
        }
      ]
    }
  });

  const floodRiskLayer = new FeatureLayer({
    url: 'https://services5.arcgis.com/VoTFsdkrXCTvwVPC/arcgis/rest/services/Flood_Risk_Area/FeatureServer/0',
    title: 'Flood Risk',
    outFields: ['*'],
    visible: false,
    renderer: {
      type: 'unique-value',
      field: 'risk_level',
      defaultSymbol: {
        type: 'simple-fill',
        color: 'rgba(200, 200, 200, 0.4)',
        outline: { color: 'rgba(100, 100, 100, 0.8)', width: 1 }
      },
      uniqueValueInfos: [
        {
          value: 'Low',
          symbol: {
            type: 'simple-fill',
            color: 'rgba(144, 238, 144, 0.5)',
            outline: { color: 'rgba(34, 139, 34, 0.8)', width: 1.5 }
          },
          label: 'Low Risk'
        },
        {
          value: 'Medium',
          symbol: {
            type: 'simple-fill',
            color: 'rgba(255, 215, 0, 0.5)',
            outline: { color: 'rgba(184, 134, 11, 0.8)', width: 1.5 }
          },
          label: 'Medium Risk'
        },
        {
          value: 'High',
          symbol: {
            type: 'simple-fill',
            color: 'rgba(255, 140, 0, 0.5)',
            outline: { color: 'rgba(255, 69, 0, 0.8)', width: 1.5 }
          },
          label: 'High Risk'
        },
        {
          value: 'Very High',
          symbol: {
            type: 'simple-fill',
            color: 'rgba(220, 20, 60, 0.6)',
            outline: { color: 'rgba(139, 0, 0, 0.9)', width: 1.5 }
          },
          label: 'Very High Risk'
        }
      ]
    },
    popupTemplate: {
      title: 'Flood Risk Area',
      content: [
        {
          type: 'fields',
          fieldInfos: [
            { fieldName: 'risk_level', label: 'Risk Level' },
            { fieldName: 'flood_frequency', label: 'Flood Frequency' },
            { fieldName: 'area_name', label: 'Area Name' },
            { fieldName: 'affected_population', label: 'Affected Population' }
          ]
        }
      ]
    }
  });

  const populationDensityLayer = new FeatureLayer({
    url: 'https://services5.arcgis.com/VoTFsdkrXCTvwVPC/arcgis/rest/services/Population_Density/FeatureServer',
    title: 'Population Density',
    outFields: ['*'],
    visible: false,
    renderer: {
      type: 'class-breaks',
      field: 'density_per_km2',
      classBreakInfos: [
        {
          minValue: 0,
          maxValue: 50,
          symbol: {
            type: 'simple-marker',
            size: 8,
            color: 'rgba(240, 240, 240, 0.8)',
            outline: { color: 'rgba(150, 150, 150, 0.8)', width: 0.5 }
          },
          label: 'Very Low (0-50)'
        },
        {
          minValue: 51,
          maxValue: 150,
          symbol: {
            type: 'simple-marker',
            size: 10,
            color: 'rgba(200, 220, 240, 0.8)',
            outline: { color: 'rgba(100, 150, 200, 0.8)', width: 0.5 }
          },
          label: 'Low (51-150)'
        },
        {
          minValue: 151,
          maxValue: 300,
          symbol: {
            type: 'simple-marker',
            size: 12,
            color: 'rgba(144, 176, 224, 0.8)',
            outline: { color: 'rgba(70, 130, 180, 0.8)', width: 0.5 }
          },
          label: 'Medium (151-300)'
        },
        {
          minValue: 301,
          maxValue: 500,
          symbol: {
            type: 'simple-marker',
            size: 14,
            color: 'rgba(70, 130, 180, 0.8)',
            outline: { color: 'rgba(25, 80, 140, 0.8)', width: 0.5 }
          },
          label: 'High (301-500)'
        },
        {
          minValue: 501,
          maxValue: 10000,
          symbol: {
            type: 'simple-marker',
            size: 16,
            color: 'rgba(25, 80, 140, 1)',
            outline: { color: 'rgba(10, 40, 100, 0.8)', width: 0.5 }
          },
          label: 'Very High (500+)'
        }
      ]
    },
    popupTemplate: {
      title: 'Population Density',
      content: [
        {
          type: 'fields',
          fieldInfos: [
            { fieldName: 'density_per_km2', label: 'Population Density (per km²)' },
            { fieldName: 'total_population', label: 'Total Population' },
            { fieldName: 'area_name', label: 'Area Name' },
            { fieldName: 'region', label: 'Region' }
          ]
        }
      ]
    }
  });

  const hospitalsLayer = new FeatureLayer({
    url: 'https://services5.arcgis.com/VoTFsdkrXCTvwVPC/arcgis/rest/services/Hospitals_Clinics/FeatureServer/0',
    title: 'Hospitals & Clinics',
    outFields: ['*'],
    visible: false,
    renderer: {
      type: 'unique-value',
      field: 'type',
      defaultSymbol: {
        type: 'simple-marker',
        size: 10,
        color: 'gray',
        outline: { color: 'white', width: 1 }
      },
      uniqueValueInfos: [
        {
          value: 'Hospital',
          symbol: {
            type: 'text',
            color: '#d7191c',
            text: 'H',
            font: { size: 16, weight: 'bold' },
            haloColor: 'white',
            haloSize: 2
          },
          label: 'Hospital'
        },
        {
          value: 'Clinic',
          symbol: {
            type: 'text',
            color: '#0099ff',
            text: 'C',
            font: { size: 16, weight: 'bold' },
            haloColor: 'white',
            haloSize: 2
          },
          label: 'Clinic'
        }
      ]
    },
    popupTemplate: {
      title: '{name}',
      content: [
        {
          type: 'fields',
          fieldInfos: [
            { fieldName: 'type', label: 'Type' },
            { fieldName: 'name', label: 'Name' },
            { fieldName: 'bed_capacity', label: 'Bed Capacity' },
            { fieldName: 'emergency_available', label: 'Emergency Available' },
            { fieldName: 'lat', label: 'Latitude' },
            { fieldName: 'lon', label: 'Longitude' }
          ]
        }
      ]
    }
  });

  const esgComplianceLayer = new FeatureLayer({
    url: 'https://services5.arcgis.com/VoTFsdkrXCTvwVPC/arcgis/rest/services/ESG_Compliance/FeatureServer/0',
    title: 'ESG Compliance',
    outFields: ['*'],
    visible: false,
    renderer: {
      type: 'class-breaks',
      field: 'overall_score',
      classBreakInfos: [
        {
          minValue: 0,
          maxValue: 59,
          symbol: {
            type: 'simple-marker',
            size: 12,
            color: 'rgba(215, 25, 28, 0.8)',
            outline: { color: 'rgba(165, 0, 38, 0.8)', width: 1 }
          },
          label: '<60'
        },
        {
          minValue: 60,
          maxValue: 79,
          symbol: {
            type: 'simple-marker',
            size: 12,
            color: 'rgba(253, 174, 97, 0.8)',
            outline: { color: 'rgba(215, 69, 0, 0.8)', width: 1 }
          },
          label: '60-79'
        },
        {
          minValue: 80,
          maxValue: 100,
          symbol: {
            type: 'simple-marker',
            size: 12,
            color: 'rgba(26, 150, 65, 0.8)',
            outline: { color: 'rgba(0, 104, 55, 0.8)', width: 1 }
          },
          label: '80-100'
        }
      ]
    },
    popupTemplate: {
      title: 'ESG Compliance',
      content: [
        {
          type: 'fields',
          fieldInfos: [
            { fieldName: 'overall_score', label: 'Overall Score' },
            { fieldName: 'entity_name', label: 'Entity' },
            { fieldName: 'assessment_date', label: 'Assessment Date' }
          ]
        }
      ]
    }
  });

  const environmentalViolationsLayer = new FeatureLayer({
    url: 'https://services5.arcgis.com/VoTFsdkrXCTvwVPC/arcgis/rest/services/Environmental_Violations/FeatureServer/0',
    title: 'Environmental Violations',
    outFields: ['*'],
    visible: false,
    renderer: {
      type: 'unique-value',
      field: 'status',
      defaultSymbol: {
        type: 'simple-marker',
        size: 10,
        color: 'rgba(150, 150, 150, 0.8)',
        outline: { color: 'white', width: 1 }
      },
      uniqueValueInfos: [
        {
          value: 'Ongoing',
          symbol: {
            type: 'simple-marker',
            size: 12,
            color: 'rgba(215, 25, 28, 0.8)',
            outline: { color: 'rgba(165, 0, 38, 0.8)', width: 1 }
          },
          label: 'Ongoing'
        },
        {
          value: 'Under Investigation',
          symbol: {
            type: 'simple-marker',
            size: 12,
            color: 'rgba(253, 174, 97, 0.8)',
            outline: { color: 'rgba(215, 69, 0, 0.8)', width: 1 }
          },
          label: 'Under Investigation'
        },
        {
          value: 'Resolved',
          symbol: {
            type: 'simple-marker',
            size: 12,
            color: 'rgba(26, 150, 65, 0.8)',
            outline: { color: 'rgba(0, 104, 55, 0.8)', width: 1 }
          },
          label: 'Resolved'
        }
      ]
    },
    popupTemplate: {
      title: 'Environmental Violation',
      content: [
        {
          type: 'fields',
          fieldInfos: [
            { fieldName: 'case_id', label: 'Case ID' },
            { fieldName: 'company', label: 'Company' },
            { fieldName: 'violation_type', label: 'Violation Type' },
            { fieldName: 'status', label: 'Status' },
            { fieldName: 'fine_amount', label: 'Fine Amount' },
            { fieldName: 'lat', label: 'Latitude' },
            { fieldName: 'lon', label: 'Longitude' }
          ]
        }
      ]
    }
  });

  map.addMany([fireHotspotsLayer, floodRiskLayer, populationDensityLayer, hospitalsLayer, esgComplianceLayer, environmentalViolationsLayer]);
  airQualityLayer.visible = false;
  industrialEmissionsLayer.visible = false;

  // Graphics layers for buffer and measurement tools (added after feature layers to be on top)
  const bufferGraphicsLayer = new GraphicsLayer();
  map.add(bufferGraphicsLayer);

  const measurementGraphicsLayer = new GraphicsLayer();
  map.add(measurementGraphicsLayer);

  // dedicated layer for sketch-analysis buffer visualization
  const sketchBufferGraphicsLayer = new GraphicsLayer({ title: 'Sketch Analysis Buffer' });
  map.add(sketchBufferGraphicsLayer);

  console.log('Graphics layers added to map');



  const view = new MapView({
    container: 'viewDiv',
    map: map,
    // extent that covers peninsular Malaysia and Sabah/Sarawak
    extent: {
      xmin: 99.0,
      ymin: -1.5,
      xmax: 119.0,
      ymax: 7.5,
      spatialReference: { wkid: 4326 }
    }
  });

  // Sketch graphics layer for user drawings
  const sketchGraphicsLayer = new GraphicsLayer();
  map.add(sketchGraphicsLayer);

  // layer to hold features loaded from Supabase
  const savedSketchesLayer = new GraphicsLayer({ title: 'Saved sketches' });
  map.add(savedSketchesLayer);
  let savedSketchRecords = [];
  let allSavedSketchRows = [];
  let savedSketchSelectEl = null;
  let projectFilterSelectEl = null;
  let projectNameInputEl = null;
  let authStatusEl = null;
  let authSignInBtnEl = null;
  let authSignUpBtnEl = null;
  let authSignOutBtnEl = null;
  let currentUser = null;
  let authStateSubscription = null;
  let lastVisibleSavedCount = 0;
  let sketchTableSupportsUserId = true;
  let sketchTableSupportsProjectName = true;
  const defaultProjectName = 'General';
  const sketchBufferStorageKey = 'esg_sketch_analysis_buffers_v1';
  let allSketchesAnalysis = [];

  // create the Sketch widget allowing polygon/point drawing
  const sketch = new Sketch({
    view: view,
    layer: sketchGraphicsLayer,
    availableCreateTools: ['polygon', 'point'],
    creationMode: 'update',
    visibleElements: {
      settingsMenu: false  // hide snapping/geometry guides/feature-to-feature
    }
  });

  view.ui.add(sketch, 'top-right');

  function getSavedSketchSymbol(geometryType) {
    if (geometryType === 'point') {
      return {
        type: 'simple-marker',
        color: [0, 150, 200, 0.6],
        size: 9,
        outline: { color: 'blue', width: 1 }
      };
    }

    return {
      type: 'simple-fill',
      color: [0, 150, 200, 0.2],
      outline: { color: 'blue', width: 1 }
    };
  }

  async function toWgs84Geometry(geometry) {
    if (!geometry || !geometry.spatialReference || geometry.spatialReference.wkid === 4326) {
      return geometry;
    }

    try {
      await projection.load();
      const projected = projection.project(geometry, { wkid: 4326 });
      return projected || geometry;
    } catch (err) {
      console.warn('Projection to WGS84 failed, using original geometry:', err);
      return geometry;
    }
  }

  function arcgisToGeojson(geometry) {
    if (!geometry) return null;

    if (geometry.type === 'polygon') {
      return {
        type: 'Polygon',
        coordinates: geometry.rings
      };
    }

    if (geometry.type === 'point') {
      return {
        type: 'Point',
        coordinates: [geometry.x ?? geometry.longitude, geometry.y ?? geometry.latitude]
      };
    }

    return null;
  }

  function normalizeStoredGeometry(rawGeom) {
    if (!rawGeom) return null;

    let geom = rawGeom;

    // If geometry is serialized text, attempt JSON parse first.
    if (typeof geom === 'string') {
      const ewkt = geom.replace(/^SRID=\d+;/i, '').trim();

      const pointMatch = ewkt.match(/^POINT\s*\(\s*(-?\d+(?:\.\d+)?)\s+(-?\d+(?:\.\d+)?)\s*\)$/i);
      if (pointMatch) {
        return {
          type: 'Point',
          coordinates: [parseFloat(pointMatch[1]), parseFloat(pointMatch[2])]
        };
      }

      const polygonMatch = ewkt.match(/^POLYGON\s*\(\((.+)\)\)$/i);
      if (polygonMatch) {
        const ring = polygonMatch[1]
          .split(',')
          .map(pair => pair.trim().split(/\s+/).map(Number));
        return { type: 'Polygon', coordinates: [ring] };
      }

      try {
        geom = JSON.parse(geom);
      } catch (_err) {
        return null;
      }
    }

    // Handle Feature wrapper format.
    if (geom.type === 'Feature' && geom.geometry) {
      geom = geom.geometry;
    }

    // Handle ArcGIS-style geometry objects.
    if (geom.type === 'point' && (geom.longitude !== undefined || geom.x !== undefined)) {
      return {
        type: 'Point',
        coordinates: [geom.longitude ?? geom.x, geom.latitude ?? geom.y]
      };
    }

    if (geom.type === 'polygon' && Array.isArray(geom.rings)) {
      return {
        type: 'Polygon',
        coordinates: geom.rings
      };
    }

    // Keep only supported GeoJSON geometry types.
    if (
      (geom.type === 'Point' || geom.type === 'Polygon') &&
      Array.isArray(geom.coordinates)
    ) {
      return geom;
    }

    return null;
  }

  function geojsonToArcgis(geojson) {
    const normalized = normalizeStoredGeometry(geojson);
    if (!normalized || !normalized.type) return null;

    if (normalized.type === 'Point' && Array.isArray(normalized.coordinates)) {
      return {
        type: 'point',
        longitude: normalized.coordinates[0],
        latitude: normalized.coordinates[1],
        spatialReference: { wkid: 4326 }
      };
    }

    if (normalized.type === 'Polygon' && Array.isArray(normalized.coordinates)) {
      return {
        type: 'polygon',
        rings: normalized.coordinates,
        spatialReference: { wkid: 4326 }
      };
    }

    return null;
  }

  function getPersistedSketchBuffers() {
    try {
      const raw = localStorage.getItem(sketchBufferStorageKey);
      if (!raw) return [];
      const parsed = JSON.parse(raw);
      return Array.isArray(parsed) ? parsed : [];
    } catch (err) {
      console.warn('Unable to read persisted analysis buffers:', err);
      return [];
    }
  }

  function setPersistedSketchBuffers(items) {
    try {
      localStorage.setItem(sketchBufferStorageKey, JSON.stringify(items));
    } catch (err) {
      console.warn('Unable to persist analysis buffers:', err);
    }
  }

  function drawSketchAnalysisBuffer(bufferGeometry, centerGeometry) {
    const bufferGraphic = new Graphic({
      geometry: bufferGeometry,
      symbol: {
        type: 'simple-fill',
        color: [23, 54, 93, 0.12],
        outline: {
          color: [23, 54, 93, 0.9],
          width: 2
        }
      }
    });

    const centerGraphic = new Graphic({
      geometry: centerGeometry,
      symbol: {
        type: 'simple-marker',
        size: 8,
        color: '#d7191c',
        outline: { color: 'white', width: 1.5 }
      }
    });

    sketchBufferGraphicsLayer.addMany([bufferGraphic, centerGraphic]);
  }

  async function persistSketchAnalysisBuffer(bufferGeometry, centerGeometry, distanceKm) {
    const wgsBuffer = await toWgs84Geometry(bufferGeometry);
    const wgsCenter = await toWgs84Geometry(centerGeometry);

    const bufferGeojson = arcgisToGeojson(wgsBuffer);
    const centerGeojson = arcgisToGeojson(wgsCenter);

    if (!bufferGeojson || !centerGeojson) {
      return;
    }

    const existing = getPersistedSketchBuffers();
    existing.push({
      id: 'buf_' + Date.now() + '_' + Math.random().toString(36).slice(2, 8),
      created_at: new Date().toISOString(),
      distance_km: distanceKm,
      buffer: bufferGeojson,
      center: centerGeojson
    });

    setPersistedSketchBuffers(existing);
  }

  function restorePersistedSketchBuffers() {
    sketchBufferGraphicsLayer.removeAll();
    const stored = getPersistedSketchBuffers();

    stored.forEach(function (item) {
      const bufferGeometry = geojsonToArcgis(item.buffer);
      const centerGeometry = geojsonToArcgis(item.center);
      if (!bufferGeometry || !centerGeometry) {
        return;
      }

      drawSketchAnalysisBuffer(bufferGeometry, centerGeometry);
    });

    if (stored.length > 0) {
      console.log('Restored', stored.length, 'persisted analysis buffer(s)');
    }
  }

  // Sketch analysis popup UI on top of the map
  let lastAnalysisResult = null;

  const analysisPopup = document.createElement('div');
  analysisPopup.style.cssText = 'position:absolute;top:16px;right:16px;width:440px;max-height:58vh;background:white;border:1px solid #dde4f0;border-radius:8px;box-shadow:0 10px 30px rgba(0,0,0,0.2);z-index:1200;overflow:hidden;display:none;';
  analysisPopup.innerHTML = `
    <div id="analysisPopupHeader" style="display:flex;justify-content:space-between;align-items:center;padding:0.75rem 1rem;background:#f4f7fb;border-bottom:1px solid #dde4f0;gap:0.5rem;">
      <h3 style="margin:0;color:#17365d;font-size:0.95rem;">Sketch Spatial Analysis</h3>
      <div style="display:flex;align-items:center;gap:0.4rem;">
        <button id="exportAnalysisCsv" style="width:auto;display:inline-flex;align-items:center;justify-content:center;padding:0.2rem 0.5rem;margin:0;background:#17365d;border:0;color:#fff;border-radius:4px;font-size:0.75rem;line-height:1.2;cursor:pointer;">CSV</button>
        <button id="exportAnalysisPdf" style="width:auto;display:inline-flex;align-items:center;justify-content:center;padding:0.2rem 0.5rem;margin:0;background:#264d86;border:0;color:#fff;border-radius:4px;font-size:0.75rem;line-height:1.2;cursor:pointer;">PDF</button>
        <button id="closeAnalysisPopup" title="Collapse analysis" aria-label="Collapse analysis" style="width:auto;display:inline-flex;align-items:center;justify-content:center;padding:0;margin:0;background:none;border:none;color:#17365d;font-size:1.1rem;line-height:1;cursor:pointer;">×</button>
      </div>
    </div>
    <div id="analysisSummary" style="padding:0.6rem 1rem;color:#314661;font-size:0.82rem;border-bottom:1px solid #eef2f8;"></div>
    <div id="analysisTableContainer" style="max-height:calc(58vh - 94px);overflow:auto;">
      <table style="width:100%;border-collapse:collapse;font-size:0.82rem;">
        <thead>
          <tr>
            <th style="text-align:left;padding:0.45rem 0.6rem;border-bottom:1px solid #dde4f0;background:#f9fbfe;color:#264d86;">Type</th>
            <th style="text-align:left;padding:0.45rem 0.6rem;border-bottom:1px solid #dde4f0;background:#f9fbfe;color:#264d86;">Metric</th>
            <th style="text-align:left;padding:0.45rem 0.6rem;border-bottom:1px solid #dde4f0;background:#f9fbfe;color:#264d86;">Value</th>
          </tr>
        </thead>
        <tbody id="analysisTableBody"></tbody>
      </table>
    </div>
  `;
  document.body.appendChild(analysisPopup);

  const analysisSummaryEl = document.getElementById('analysisSummary');
  const analysisTableBodyEl = document.getElementById('analysisTableBody');
  const analysisTableContainerEl = document.getElementById('analysisTableContainer');
  const analysisPopupHeaderEl = document.getElementById('analysisPopupHeader');
  const exportAnalysisCsvBtn = document.getElementById('exportAnalysisCsv');
  const exportAnalysisPdfBtn = document.getElementById('exportAnalysisPdf');
  const closeAnalysisPopupBtn = document.getElementById('closeAnalysisPopup');
  let isAnalysisPopupCollapsed = false;
  const analysisPopupDragState = { manuallyPositioned: false, left: 0, top: 0 };

  const allAnalysisPopup = document.createElement('div');
  allAnalysisPopup.style.cssText = 'position:absolute;top:16px;right:468px;width:560px;max-height:58vh;background:white;border:1px solid #dde4f0;border-radius:8px;box-shadow:0 10px 30px rgba(0,0,0,0.2);z-index:1199;overflow:hidden;display:none;';
  allAnalysisPopup.innerHTML = `
    <div id="allAnalysisPopupHeader" style="display:flex;justify-content:space-between;align-items:center;padding:0.75rem 1rem;background:#f4f7fb;border-bottom:1px solid #dde4f0;gap:0.5rem;">
      <h3 style="margin:0;color:#17365d;font-size:0.95rem;">All Sketches Analysis</h3>
      <div style="display:flex;align-items:center;gap:0.4rem;">
        <button id="refreshAllAnalysis" style="width:auto;display:inline-flex;align-items:center;justify-content:center;padding:0.2rem 0.5rem;margin:0;background:#17365d;border:0;color:#fff;border-radius:4px;font-size:0.75rem;line-height:1.2;cursor:pointer;">Refresh</button>
        <button id="exportAllAnalysisCsv" style="width:auto;display:inline-flex;align-items:center;justify-content:center;padding:0.2rem 0.5rem;margin:0;background:#17365d;border:0;color:#fff;border-radius:4px;font-size:0.75rem;line-height:1.2;cursor:pointer;">CSV</button>
        <button id="exportAllAnalysisPdf" style="width:auto;display:inline-flex;align-items:center;justify-content:center;padding:0.2rem 0.5rem;margin:0;background:#264d86;border:0;color:#fff;border-radius:4px;font-size:0.75rem;line-height:1.2;cursor:pointer;">PDF</button>
        <button id="closeAllAnalysisPopup" title="Collapse all-sketch analysis" aria-label="Collapse all-sketch analysis" style="width:auto;display:inline-flex;align-items:center;justify-content:center;padding:0;margin:0;background:none;border:none;color:#17365d;font-size:1.1rem;line-height:1;cursor:pointer;">×</button>
      </div>
    </div>
    <div id="allAnalysisSummary" style="padding:0.6rem 1rem;color:#314661;font-size:0.82rem;border-bottom:1px solid #eef2f8;"></div>
    <div id="allAnalysisControls" style="display:flex;align-items:center;gap:0.4rem;padding:0.45rem 1rem;border-bottom:1px solid #eef2f8;background:#fcfdff;">
      <label for="allAnalysisTypeFilter" style="font-size:0.76rem;color:#264d86;font-weight:600;">Type</label>
      <select id="allAnalysisTypeFilter" style="font-size:0.76rem;padding:0.2rem 0.3rem;border:1px solid #c8d3e6;border-radius:4px;background:white;color:#17365d;">
        <option value="all">All</option>
        <option value="point">Point</option>
        <option value="polygon">Polygon</option>
      </select>
      <label for="allAnalysisSortField" style="font-size:0.76rem;color:#264d86;font-weight:600;margin-left:0.35rem;">Sort</label>
      <select id="allAnalysisSortField" style="font-size:0.76rem;padding:0.2rem 0.3rem;border:1px solid #c8d3e6;border-radius:4px;background:white;color:#17365d;">
        <option value="time">Time</option>
        <option value="intersections">Intersections</option>
        <option value="hotspots">Hotspots</option>
        <option value="stations">Stations</option>
        <option value="avgAQI">Avg AQI</option>
        <option value="maxAQI">Max AQI</option>
        <option value="bufferKm">Buffer km</option>
      </select>
      <select id="allAnalysisSortDirection" style="font-size:0.76rem;padding:0.2rem 0.3rem;border:1px solid #c8d3e6;border-radius:4px;background:white;color:#17365d;">
        <option value="desc">Desc</option>
        <option value="asc">Asc</option>
      </select>
    </div>
    <div id="allAnalysisTableContainer" style="max-height:calc(58vh - 138px);overflow:auto;">
      <table style="width:100%;border-collapse:collapse;font-size:0.8rem;">
        <thead>
          <tr>
            <th style="text-align:left;padding:0.45rem 0.5rem;border-bottom:1px solid #dde4f0;background:#f9fbfe;color:#264d86;">#</th>
            <th style="text-align:left;padding:0.45rem 0.5rem;border-bottom:1px solid #dde4f0;background:#f9fbfe;color:#264d86;">Saved ID</th>
            <th style="text-align:left;padding:0.45rem 0.5rem;border-bottom:1px solid #dde4f0;background:#f9fbfe;color:#264d86;">Time</th>
            <th style="text-align:left;padding:0.45rem 0.5rem;border-bottom:1px solid #dde4f0;background:#f9fbfe;color:#264d86;">Type</th>
            <th style="text-align:left;padding:0.45rem 0.5rem;border-bottom:1px solid #dde4f0;background:#f9fbfe;color:#264d86;">Intersections</th>
            <th style="text-align:left;padding:0.45rem 0.5rem;border-bottom:1px solid #dde4f0;background:#f9fbfe;color:#264d86;">Hotspots</th>
            <th style="text-align:left;padding:0.45rem 0.5rem;border-bottom:1px solid #dde4f0;background:#f9fbfe;color:#264d86;">Stations</th>
            <th style="text-align:left;padding:0.45rem 0.5rem;border-bottom:1px solid #dde4f0;background:#f9fbfe;color:#264d86;">Avg AQI</th>
            <th style="text-align:left;padding:0.45rem 0.5rem;border-bottom:1px solid #dde4f0;background:#f9fbfe;color:#264d86;">Max AQI</th>
            <th style="text-align:left;padding:0.45rem 0.5rem;border-bottom:1px solid #dde4f0;background:#f9fbfe;color:#264d86;">Buffer km</th>
          </tr>
        </thead>
        <tbody id="allAnalysisTableBody"></tbody>
      </table>
    </div>
  `;
  document.body.appendChild(allAnalysisPopup);

  const allAnalysisSummaryEl = document.getElementById('allAnalysisSummary');
  const allAnalysisPopupHeaderEl = document.getElementById('allAnalysisPopupHeader');
  const allAnalysisControlsEl = document.getElementById('allAnalysisControls');
  const allAnalysisTableBodyEl = document.getElementById('allAnalysisTableBody');
  const allAnalysisTableContainerEl = document.getElementById('allAnalysisTableContainer');
  const closeAllAnalysisPopupBtn = document.getElementById('closeAllAnalysisPopup');
  const refreshAllAnalysisBtn = document.getElementById('refreshAllAnalysis');
  const exportAllAnalysisCsvBtn = document.getElementById('exportAllAnalysisCsv');
  const exportAllAnalysisPdfBtn = document.getElementById('exportAllAnalysisPdf');
  const allAnalysisTypeFilterEl = document.getElementById('allAnalysisTypeFilter');
  const allAnalysisSortFieldEl = document.getElementById('allAnalysisSortField');
  const allAnalysisSortDirectionEl = document.getElementById('allAnalysisSortDirection');
  let isAllAnalysisPopupCollapsed = false;
  const allAnalysisPopupDragState = { manuallyPositioned: false, left: 0, top: 0 };

  function setAllAnalysisPopupCollapsed(collapsed) {
    isAllAnalysisPopupCollapsed = Boolean(collapsed);

    allAnalysisSummaryEl.style.display = isAllAnalysisPopupCollapsed ? 'none' : 'block';
    allAnalysisControlsEl.style.display = isAllAnalysisPopupCollapsed ? 'none' : 'flex';
    allAnalysisTableContainerEl.style.display = isAllAnalysisPopupCollapsed ? 'none' : 'block';

    closeAllAnalysisPopupBtn.textContent = isAllAnalysisPopupCollapsed ? '+' : '×';
    closeAllAnalysisPopupBtn.title = isAllAnalysisPopupCollapsed ? 'Expand all-sketch analysis' : 'Collapse all-sketch analysis';
    closeAllAnalysisPopupBtn.setAttribute('aria-label', isAllAnalysisPopupCollapsed ? 'Expand all-sketch analysis' : 'Collapse all-sketch analysis');

    allAnalysisPopupHeaderEl.style.borderBottom = isAllAnalysisPopupCollapsed ? '0' : '1px solid #dde4f0';
  }

  closeAllAnalysisPopupBtn.addEventListener('click', function () {
    setAllAnalysisPopupCollapsed(!isAllAnalysisPopupCollapsed);
    allAnalysisPopup.style.display = 'block';
  });

  refreshAllAnalysisBtn.addEventListener('click', function () {
    Promise.resolve(openAllSketchesAnalysisPopup()).catch(function (err) {
      console.error('All-sketches analysis refresh error:', err);
    });
  });

  exportAllAnalysisCsvBtn.addEventListener('click', function () {
    try {
      exportAllSketchesAnalysisCsv();
    } catch (err) {
      console.error('All-sketches CSV export error:', err);
    }
  });

  exportAllAnalysisPdfBtn.addEventListener('click', function () {
    try {
      exportAllSketchesAnalysisPdf();
    } catch (err) {
      console.error('All-sketches PDF export error:', err);
    }
  });

  [allAnalysisTypeFilterEl, allAnalysisSortFieldEl, allAnalysisSortDirectionEl].forEach(function (el) {
    el.addEventListener('change', function () {
      renderAllSketchesAnalysisTable();
    });
  });

  function setAnalysisPopupCollapsed(collapsed) {
    isAnalysisPopupCollapsed = Boolean(collapsed);

    analysisSummaryEl.style.display = isAnalysisPopupCollapsed ? 'none' : 'block';
    analysisTableContainerEl.style.display = isAnalysisPopupCollapsed ? 'none' : 'block';

    closeAnalysisPopupBtn.textContent = isAnalysisPopupCollapsed ? '+' : '×';
    closeAnalysisPopupBtn.title = isAnalysisPopupCollapsed ? 'Expand analysis' : 'Collapse analysis';
    closeAnalysisPopupBtn.setAttribute('aria-label', isAnalysisPopupCollapsed ? 'Expand analysis' : 'Collapse analysis');

    analysisPopupHeaderEl.style.borderBottom = isAnalysisPopupCollapsed ? '0' : '1px solid #dde4f0';
  }

  closeAnalysisPopupBtn.addEventListener('click', function () {
    setAnalysisPopupCollapsed(!isAnalysisPopupCollapsed);
    analysisPopup.style.display = 'block';
  });

  function escapeHtml(value) {
    return String(value)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');
  }

  function createBadge(label, backgroundColor, textColor) {
    return '<span style="display:inline-block;padding:0.08rem 0.42rem;border-radius:999px;background:' + backgroundColor + ';color:' + textColor + ';font-size:0.72rem;font-weight:700;margin-left:0.35rem;">' + escapeHtml(label) + '</span>';
  }

  function getAqiCategory(aqi) {
    if (aqi <= 50) return { label: 'Good', bg: '#e6f6ea', fg: '#0f7a30' };
    if (aqi <= 100) return { label: 'Moderate', bg: '#fff7e5', fg: '#8a5a00' };
    if (aqi <= 200) return { label: 'Poor', bg: '#ffeddc', fg: '#9a3b00' };
    return { label: 'Very Poor', bg: '#ffe4e4', fg: '#a51d1d' };
  }

  function getHotspotSeverity(brightness) {
    const value = Number(brightness);
    if (!Number.isFinite(value)) return { label: 'Unknown', bg: '#eef2f8', fg: '#314661' };
    if (value >= 350) return { label: 'Extreme', bg: '#fde8e8', fg: '#9b1c1c' };
    if (value >= 300) return { label: 'High', bg: '#ffeddc', fg: '#9a3b00' };
    if (value >= 200) return { label: 'Medium', bg: '#fff7e5', fg: '#8a5a00' };
    return { label: 'Low', bg: '#e6f6ea', fg: '#0f7a30' };
  }

  function getTopRightWidgetsBottom() {
    const container = view && view.container;
    if (!container) {
      return 44;
    }

    const topRight = container.querySelector('.esri-ui-top-right');
    if (!topRight) {
      return 44;
    }

    let maxBottom = Math.ceil(topRight.getBoundingClientRect().top || 44);
    const components = topRight.querySelectorAll('.esri-component');

    components.forEach(function (node) {
      const rect = node.getBoundingClientRect();
      if (rect.width <= 0 || rect.height <= 0) {
        return;
      }

      const computed = window.getComputedStyle(node);
      if (computed.display === 'none' || computed.visibility === 'hidden') {
        return;
      }

      maxBottom = Math.max(maxBottom, Math.ceil(rect.bottom));
    });

    return maxBottom;
  }

  function clampToRange(value, min, max) {
    return Math.min(Math.max(value, min), max);
  }

  function getPopupConstraintBounds(popupEl, minTop) {
    const rect = popupEl.getBoundingClientRect();
    const width = rect.width || popupEl.offsetWidth || 320;
    const height = rect.height || popupEl.offsetHeight || 180;

    return {
      minLeft: 8,
      maxLeft: Math.max(8, window.innerWidth - width - 8),
      minTop: minTop,
      maxTop: Math.max(minTop, window.innerHeight - height - 8)
    };
  }

  function applyManualPopupPosition(popupEl, dragState, minTop) {
    const bounds = getPopupConstraintBounds(popupEl, minTop);
    dragState.left = clampToRange(dragState.left, bounds.minLeft, bounds.maxLeft);
    dragState.top = clampToRange(dragState.top, bounds.minTop, bounds.maxTop);

    popupEl.style.left = dragState.left + 'px';
    popupEl.style.top = dragState.top + 'px';
    popupEl.style.right = '';
  }

  function makePopupDraggable(popupEl, headerEl, dragState, getMinTop) {
    if (!popupEl || !headerEl) {
      return;
    }

    headerEl.style.cursor = 'move';
    headerEl.style.touchAction = 'none';

    headerEl.addEventListener('pointerdown', function (event) {
      if (event.pointerType === 'mouse' && event.button !== 0) {
        return;
      }

      const target = event.target;
      if (target && typeof target.closest === 'function' && target.closest('button,select,input,textarea,a')) {
        return;
      }

      const startRect = popupEl.getBoundingClientRect();
      const popupWidth = startRect.width || popupEl.offsetWidth || 320;
      const popupHeight = startRect.height || popupEl.offsetHeight || 180;
      const offsetX = event.clientX - startRect.left;
      const offsetY = event.clientY - startRect.top;

      dragState.manuallyPositioned = true;
      dragState.left = startRect.left;
      dragState.top = startRect.top;
      popupEl.style.left = dragState.left + 'px';
      popupEl.style.top = dragState.top + 'px';
      popupEl.style.right = '';

      if (typeof headerEl.setPointerCapture === 'function') {
        try {
          headerEl.setPointerCapture(event.pointerId);
        } catch (err) {
          // Ignore capture failures.
        }
      }

      function moveTo(clientX, clientY) {
        const minTop = getMinTop();
        const maxLeft = Math.max(8, window.innerWidth - popupWidth - 8);
        const maxTop = Math.max(minTop, window.innerHeight - popupHeight - 8);

        dragState.left = clampToRange(clientX - offsetX, 8, maxLeft);
        dragState.top = clampToRange(clientY - offsetY, minTop, maxTop);
        popupEl.style.left = dragState.left + 'px';
        popupEl.style.top = dragState.top + 'px';
        popupEl.style.right = '';
      }

      function onPointerMove(moveEvent) {
        moveTo(moveEvent.clientX, moveEvent.clientY);
        moveEvent.preventDefault();
      }

      function cleanup(upEvent) {
        headerEl.removeEventListener('pointermove', onPointerMove);
        headerEl.removeEventListener('pointerup', onPointerUp);
        headerEl.removeEventListener('pointercancel', onPointerCancel);

        if (typeof headerEl.releasePointerCapture === 'function' && upEvent) {
          try {
            headerEl.releasePointerCapture(upEvent.pointerId);
          } catch (err) {
            // Ignore release failures.
          }
        }
      }

      function onPointerUp(upEvent) {
        cleanup(upEvent);
      }

      function onPointerCancel(cancelEvent) {
        cleanup(cancelEvent);
      }

      headerEl.addEventListener('pointermove', onPointerMove);
      headerEl.addEventListener('pointerup', onPointerUp);
      headerEl.addEventListener('pointercancel', onPointerCancel);
    });
  }

  function positionAnalysisPopup() {
    const isMobile = window.innerWidth <= 900;
    if (!isMobile) {
      const widgetsBottom = getTopRightWidgetsBottom();
      const top = Math.max(16, widgetsBottom + 8);
      analysisPopup.style.width = '440px';

      if (analysisPopupDragState.manuallyPositioned) {
        applyManualPopupPosition(analysisPopup, analysisPopupDragState, top);
        const availableHeight = Math.max(220, window.innerHeight - analysisPopupDragState.top - 16);
        analysisPopup.style.maxHeight = availableHeight + 'px';
        analysisTableContainerEl.style.maxHeight = Math.max(140, availableHeight - 94) + 'px';
      } else {
        const availableHeight = Math.max(220, window.innerHeight - top - 16);
        analysisPopup.style.top = top + 'px';
        analysisPopup.style.right = '16px';
        analysisPopup.style.left = '';
        analysisPopup.style.maxHeight = availableHeight + 'px';
        analysisTableContainerEl.style.maxHeight = Math.max(140, availableHeight - 94) + 'px';
      }

      return;
    }

    analysisPopupDragState.manuallyPositioned = false;

    const sketchNode = document.querySelector('.esri-sketch');
    const widgetBottom = sketchNode ? Math.ceil(sketchNode.getBoundingClientRect().bottom) : 44;
    const top = Math.max(8, widgetBottom + 8);
    const availableHeight = Math.max(220, window.innerHeight - top - 8);

    analysisPopup.style.left = '8px';
    analysisPopup.style.right = '8px';
    analysisPopup.style.width = 'auto';
    analysisPopup.style.top = top + 'px';
    analysisPopup.style.maxHeight = availableHeight + 'px';
    analysisTableContainerEl.style.maxHeight = Math.max(140, availableHeight - 94) + 'px';
  }

  function positionAllAnalysisPopup() {
    const isMobile = window.innerWidth <= 900;
    if (!isMobile) {
      allAnalysisPopup.style.width = '560px';

      if (allAnalysisPopupDragState.manuallyPositioned) {
        applyManualPopupPosition(allAnalysisPopup, allAnalysisPopupDragState, 16);
        const availableHeight = Math.max(220, window.innerHeight - allAnalysisPopupDragState.top - 16);
        allAnalysisPopup.style.maxHeight = availableHeight + 'px';
        allAnalysisTableContainerEl.style.maxHeight = Math.max(100, availableHeight - 138) + 'px';
      } else {
        allAnalysisPopup.style.top = '16px';
        allAnalysisPopup.style.right = '468px';
        allAnalysisPopup.style.left = '';
        allAnalysisPopup.style.maxHeight = '58vh';
        allAnalysisTableContainerEl.style.maxHeight = 'calc(58vh - 138px)';
      }

      return;
    }

    allAnalysisPopupDragState.manuallyPositioned = false;

    const sketchNode = document.querySelector('.esri-sketch');
    const widgetBottom = sketchNode ? Math.ceil(sketchNode.getBoundingClientRect().bottom) : 44;
    const top = Math.max(8, widgetBottom + 8);
    const availableHeight = Math.max(220, window.innerHeight - top - 8);

    allAnalysisPopup.style.left = '8px';
    allAnalysisPopup.style.right = '8px';
    allAnalysisPopup.style.width = 'auto';
    allAnalysisPopup.style.top = top + 'px';
    allAnalysisPopup.style.maxHeight = availableHeight + 'px';
    allAnalysisTableContainerEl.style.maxHeight = Math.max(100, availableHeight - 138) + 'px';
  }

  makePopupDraggable(analysisPopup, analysisPopupHeaderEl, analysisPopupDragState, function () {
    return window.innerWidth <= 900 ? 8 : Math.max(16, getTopRightWidgetsBottom() + 8);
  });

  makePopupDraggable(allAnalysisPopup, allAnalysisPopupHeaderEl, allAnalysisPopupDragState, function () {
    return window.innerWidth <= 900 ? 8 : 16;
  });

  window.addEventListener('resize', positionAnalysisPopup);
  window.addEventListener('orientationchange', positionAnalysisPopup);
  window.addEventListener('resize', positionAllAnalysisPopup);
  window.addEventListener('orientationchange', positionAllAnalysisPopup);

  async function buildAllSketchesAnalysisFromSavedSketches() {
    if (savedSketchRecords.length === 0) {
      await loadSketches();
    }

    const sourceRecords = savedSketchRecords.slice();
    if (sourceRecords.length === 0) {
      allSketchesAnalysis = [];
      return;
    }

    const results = [];
    for (const record of sourceRecords) {
      const geometry = record.graphic && record.graphic.geometry;
      if (!geometry) {
        continue;
      }

      try {
        const intersections = await queryIntersections(geometry);
        const hotspots = await queryHotspotDetection(geometry);
        const pollution = await queryPollutionAnalysis(geometry);

        let bufferResults = null;
        if (geometry.type === 'point') {
          bufferResults = await queryBufferAnalysis(geometry);
        }

        results.push({
          id: record.id,
          created_at: record.created_at,
          geometryType: record.geometry_type || geometry.type,
          intersectionTotal: intersections.reduce(function (sum, item) {
            return sum + item.count;
          }, 0),
          hotspotCount: hotspots.count,
          pollutionStations: pollution.count,
          avgAQI: pollution.averageAQI,
          maxAQI: pollution.maxAQI,
          bufferKm: bufferResults ? bufferResults.distanceKm : null,
          analysisError: false
        });
      } catch (err) {
        console.error('All-sketch analysis failed for sketch ' + (record.id || 'unknown') + ':', err);
        results.push({
          id: record.id,
          created_at: record.created_at,
          geometryType: record.geometry_type || geometry.type,
          intersectionTotal: 0,
          hotspotCount: 0,
          pollutionStations: 0,
          avgAQI: 0,
          maxAQI: 0,
          bufferKm: geometry.type === 'point' ? getConfiguredBufferDistanceKm() : null,
          analysisError: true
        });
      }
    }

    allSketchesAnalysis = results;
  }

  function getAllSketchesAnalysisRows() {
    const filterType = allAnalysisTypeFilterEl ? allAnalysisTypeFilterEl.value : 'all';
    const sortField = allAnalysisSortFieldEl ? allAnalysisSortFieldEl.value : 'time';
    const sortDirection = allAnalysisSortDirectionEl ? allAnalysisSortDirectionEl.value : 'desc';

    let rows = allSketchesAnalysis.slice();
    if (filterType !== 'all') {
      rows = rows.filter(function (entry) {
        return String(entry.geometryType || '').toLowerCase() === filterType;
      });
    }

    function numValue(value, fallback) {
      const n = Number(value);
      return Number.isFinite(n) ? n : fallback;
    }

    rows.sort(function (a, b) {
      let av;
      let bv;

      if (sortField === 'time') {
        av = new Date(a.created_at || 0).getTime();
        bv = new Date(b.created_at || 0).getTime();
      } else if (sortField === 'intersections') {
        av = numValue(a.intersectionTotal, 0);
        bv = numValue(b.intersectionTotal, 0);
      } else if (sortField === 'hotspots') {
        av = numValue(a.hotspotCount, 0);
        bv = numValue(b.hotspotCount, 0);
      } else if (sortField === 'stations') {
        av = numValue(a.pollutionStations, 0);
        bv = numValue(b.pollutionStations, 0);
      } else if (sortField === 'avgAQI') {
        av = numValue(a.avgAQI, 0);
        bv = numValue(b.avgAQI, 0);
      } else if (sortField === 'maxAQI') {
        av = numValue(a.maxAQI, 0);
        bv = numValue(b.maxAQI, 0);
      } else if (sortField === 'bufferKm') {
        av = numValue(a.bufferKm, -1);
        bv = numValue(b.bufferKm, -1);
      } else {
        av = new Date(a.created_at || 0).getTime();
        bv = new Date(b.created_at || 0).getTime();
      }

      const cmp = av === bv ? 0 : (av > bv ? 1 : -1);
      return sortDirection === 'asc' ? cmp : -cmp;
    });

    return rows;
  }

  function renderAllSketchesAnalysisTable() {
    allAnalysisTableBodyEl.innerHTML = '';

    if (allSketchesAnalysis.length === 0) {
      const row = document.createElement('tr');
      const cell = document.createElement('td');
      cell.colSpan = 10;
      cell.textContent = 'No saved sketches are visible for analysis.';
      cell.style.cssText = 'padding:0.6rem;border-bottom:1px solid #eef2f8;color:#5b6d84;';
      row.appendChild(cell);
      allAnalysisTableBodyEl.appendChild(row);
      allAnalysisSummaryEl.textContent = 'Total analyses: 0';
      return;
    }

    const ordered = getAllSketchesAnalysisRows();

    if (ordered.length === 0) {
      const row = document.createElement('tr');
      const cell = document.createElement('td');
      cell.colSpan = 10;
      cell.textContent = 'No rows match the selected filter.';
      cell.style.cssText = 'padding:0.6rem;border-bottom:1px solid #eef2f8;color:#5b6d84;';
      row.appendChild(cell);
      allAnalysisTableBodyEl.appendChild(row);
      allAnalysisSummaryEl.textContent = 'Total analyses: ' + allSketchesAnalysis.length + ', Showing: 0';
      return;
    }

    let totalIntersections = 0;
    let totalHotspots = 0;
    let totalStations = 0;

    ordered.forEach(function (entry, index) {
      const row = document.createElement('tr');

      function addCell(value) {
        const cell = document.createElement('td');
        cell.textContent = value;
        cell.style.cssText = 'padding:0.42rem 0.5rem;border-bottom:1px solid #eef2f8;color:#17365d;vertical-align:top;';
        row.appendChild(cell);
      }

      let createdText = entry.created_at || '';
      if (createdText) {
        const dt = new Date(createdText);
        createdText = Number.isNaN(dt.getTime()) ? createdText : dt.toLocaleString();
      }

      totalIntersections += Number(entry.intersectionTotal || 0);
      totalHotspots += Number(entry.hotspotCount || 0);
      totalStations += Number(entry.pollutionStations || 0);

      addCell(String(index + 1));
      addCell(entry.id ? String(entry.id).slice(0, 8) : 'n/a');
      addCell(createdText || 'n/a');
      addCell(entry.geometryType || 'n/a');
      addCell(String(entry.intersectionTotal ?? 0));
      addCell(String(entry.hotspotCount ?? 0));
      addCell(String(entry.pollutionStations ?? 0));
      addCell(Number.isFinite(entry.avgAQI) ? Number(entry.avgAQI).toFixed(2) : '0.00');
      addCell(Number.isFinite(entry.maxAQI) ? Number(entry.maxAQI).toFixed(2) : '0.00');
      addCell(entry.bufferKm == null ? '-' : String(entry.bufferKm));

      allAnalysisTableBodyEl.appendChild(row);
    });

    const filterText = allAnalysisTypeFilterEl ? allAnalysisTypeFilterEl.value : 'all';
    allAnalysisSummaryEl.textContent =
      'Total analyses: ' + allSketchesAnalysis.length +
      ', Showing: ' + ordered.length +
      ', Filter: ' + filterText +
      ', Intersections: ' + totalIntersections +
      ', Hotspots: ' + totalHotspots +
      ', Stations: ' + totalStations;
  }

  function exportAllSketchesAnalysisCsv() {
    const rows = getAllSketchesAnalysisRows();
    if (rows.length === 0) {
      alert('No analysis rows available for current filter.');
      return;
    }

    const csvRows = [
      ['Index', 'Saved ID', 'Created At', 'Type', 'Intersections', 'Hotspots', 'Stations', 'Avg AQI', 'Max AQI', 'Buffer km']
    ];

    rows.forEach(function (entry, index) {
      csvRows.push([
        String(index + 1),
        entry.id ? String(entry.id) : '',
        entry.created_at || '',
        entry.geometryType || '',
        String(entry.intersectionTotal ?? 0),
        String(entry.hotspotCount ?? 0),
        String(entry.pollutionStations ?? 0),
        Number.isFinite(entry.avgAQI) ? Number(entry.avgAQI).toFixed(2) : '0.00',
        Number.isFinite(entry.maxAQI) ? Number(entry.maxAQI).toFixed(2) : '0.00',
        entry.bufferKm == null ? '' : String(entry.bufferKm)
      ]);
    });

    const csv = csvRows
      .map(function (row) {
        return row
          .map(function (cell) {
            const text = String(cell).replace(/"/g, '""');
            return '"' + text + '"';
          })
          .join(',');
      })
      .join('\n');

    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = 'all-sketches-analysis.csv';
    link.click();
    URL.revokeObjectURL(url);
  }

  async function openAllSketchesAnalysisPopup() {
    positionAllAnalysisPopup();
    allAnalysisPopup.style.display = 'block';
    setAllAnalysisPopupCollapsed(false);

    allAnalysisSummaryEl.textContent = 'Running analysis for all visible saved sketches...';
    allAnalysisTableBodyEl.innerHTML = '<tr><td colspan="10" style="padding:0.6rem;border-bottom:1px solid #eef2f8;color:#5b6d84;">Loading...</td></tr>';

    try {
      await buildAllSketchesAnalysisFromSavedSketches();
      renderAllSketchesAnalysisTable();
    } catch (err) {
      console.error('Failed to build all-sketch analysis:', err);
      allAnalysisSummaryEl.textContent = 'Analysis failed due to an unexpected error.';
      allAnalysisTableBodyEl.innerHTML = '<tr><td colspan="10" style="padding:0.6rem;border-bottom:1px solid #eef2f8;color:#b42318;">Analysis failed. Please try again after reloading sketches.</td></tr>';
    }
  }

  function addAnalysisRow(type, metric, value, badge) {
    const row = document.createElement('tr');

    const typeCell = document.createElement('td');
    typeCell.textContent = type;
    typeCell.style.cssText = 'padding:0.45rem 0.6rem;border-bottom:1px solid #eef2f8;color:#17365d;font-weight:600;vertical-align:top;';

    const metricCell = document.createElement('td');
    metricCell.textContent = metric;
    metricCell.style.cssText = 'padding:0.45rem 0.6rem;border-bottom:1px solid #eef2f8;color:#314661;vertical-align:top;';

    const valueCell = document.createElement('td');
    valueCell.innerHTML = '<span>' + escapeHtml(value) + '</span>' + (badge || '');
    valueCell.style.cssText = 'padding:0.45rem 0.6rem;border-bottom:1px solid #eef2f8;color:#17365d;vertical-align:top;';

    row.appendChild(typeCell);
    row.appendChild(metricCell);
    row.appendChild(valueCell);
    analysisTableBodyEl.appendChild(row);
  }

  function analysisRowsToCsv(result) {
    const rows = [];
    rows.push(['Type', 'Metric', 'Value']);

    rows.push(['Sketch', 'Geometry type', result.geometryType]);
    if (result.intersections.length === 0) {
      rows.push(['Intersection', 'Visible selected layers', 'None']);
    } else {
      result.intersections.forEach(function (item) {
        rows.push(['Intersection', item.layerName, item.count + ' feature(s)']);
      });
    }

    const hotspotTopBrightness = result.hotspots.items.reduce(function (maxValue, item) {
      const v = Number(item.brightness);
      return Number.isFinite(v) && v > maxValue ? v : maxValue;
    }, -Infinity);
    const hotspotSeverity = getHotspotSeverity(hotspotTopBrightness);
    rows.push(['Hotspot', 'Detected hotspots', result.hotspots.count + ' hotspot(s) - ' + hotspotSeverity.label]);

    result.hotspots.items.slice(0, 3).forEach(function (item, index) {
      rows.push([
        'Hotspot #' + (index + 1),
        'Details',
        'Brightness ' + (item.brightness ?? 'n/a') + ', FRP ' + (item.radiativePower ?? 'n/a') + ', Confidence ' + (item.confidence ?? 'n/a')
      ]);
    });

    rows.push(['Pollution', 'Stations in area', String(result.pollution.count)]);
    if (result.pollution.count > 0) {
      const avgCategory = getAqiCategory(result.pollution.averageAQI);
      const maxCategory = getAqiCategory(result.pollution.maxAQI);
      rows.push(['Pollution', 'Average AQI', result.pollution.averageAQI.toFixed(2) + ' (' + avgCategory.label + ')']);
      rows.push(['Pollution', 'Max AQI', String(result.pollution.maxAQI) + ' (' + maxCategory.label + ')']);
    }

    if (result.bufferResults) {
      rows.push(['Buffer', 'Distance', result.bufferResults.distanceKm + ' km']);
      result.bufferResults.layers.forEach(function (item) {
        rows.push(['Buffer', item.layerName, item.count + ' feature(s)']);
      });
    }

    return rows;
  }

  function getJsPdfConstructor() {
    if (window.jspdf && window.jspdf.jsPDF) {
      return window.jspdf.jsPDF;
    }

    return null;
  }

  function createPdfReportWriter(title) {
    const JsPdf = getJsPdfConstructor();
    if (!JsPdf) {
      alert('PDF library is not available. Reload the page and try again.');
      return null;
    }

    const doc = new JsPdf({ orientation: 'portrait', unit: 'pt', format: 'a4' });
    const pageWidth = doc.internal.pageSize.getWidth();
    const pageHeight = doc.internal.pageSize.getHeight();
    const marginX = 40;
    const usableWidth = pageWidth - marginX * 2;
    const bottomLimit = pageHeight - 34;
    let y = 44;

    doc.setFont('helvetica', 'bold');
    doc.setFontSize(14);
    doc.text(title, marginX, y);
    y += 18;

    doc.setFont('helvetica', 'normal');
    doc.setFontSize(9);
    doc.text('Generated: ' + new Date().toLocaleString(), marginX, y);
    y += 14;

    function writeLine(text, bold) {
      const value = String(text == null ? '' : text);
      const lines = value ? doc.splitTextToSize(value, usableWidth) : [''];

      lines.forEach(function (line) {
        if (y > bottomLimit) {
          doc.addPage();
          y = 40;
        }

        doc.setFont('helvetica', bold ? 'bold' : 'normal');
        doc.setFontSize(10);
        doc.text(line, marginX, y);
        y += 13;
      });
    }

    return {
      doc: doc,
      writeLine: writeLine,
      save: function (fileName) {
        doc.save(fileName);
      }
    };
  }

  function exportCurrentSketchAnalysisPdf() {
    if (!lastAnalysisResult) {
      alert('No sketch analysis is available yet.');
      return;
    }

    const writer = createPdfReportWriter('Sketch Spatial Analysis Report');
    if (!writer) {
      return;
    }

    writer.writeLine('Project for new sketches: ' + getCurrentProjectNameForSave());
    writer.writeLine('Summary: ' + (analysisSummaryEl.textContent || 'n/a'));
    writer.writeLine('', false);
    writer.writeLine('Type | Metric | Value', true);

    const rows = analysisRowsToCsv(lastAnalysisResult);
    rows.slice(1).forEach(function (row) {
      writer.writeLine(row[0] + ' | ' + row[1] + ' | ' + row[2]);
    });

    writer.save('sketch-spatial-analysis-report.pdf');
  }

  function exportAllSketchesAnalysisPdf() {
    const rows = getAllSketchesAnalysisRows();
    if (rows.length === 0) {
      alert('No all-sketch analysis rows available for current filter.');
      return;
    }

    const writer = createPdfReportWriter('All Sketches Analysis Report');
    if (!writer) {
      return;
    }

    const userLabel = currentUser ? (currentUser.email || currentUser.phone || currentUser.id) : 'not signed in';
    const projectFilter = getSelectedProjectFilter();
    writer.writeLine('User: ' + userLabel);
    writer.writeLine('Project filter: ' + projectFilter);
    writer.writeLine('Summary: ' + (allAnalysisSummaryEl.textContent || 'n/a'));
    writer.writeLine('', false);
    writer.writeLine('# | Saved ID | Type | Intersections | Hotspots | Stations | Avg AQI | Max AQI | Buffer km', true);

    rows.forEach(function (entry, index) {
      writer.writeLine(
        (index + 1) + ' | ' +
        (entry.id ? String(entry.id).slice(0, 8) : 'n/a') + ' | ' +
        (entry.geometryType || 'n/a') + ' | ' +
        String(entry.intersectionTotal ?? 0) + ' | ' +
        String(entry.hotspotCount ?? 0) + ' | ' +
        String(entry.pollutionStations ?? 0) + ' | ' +
        (Number.isFinite(entry.avgAQI) ? Number(entry.avgAQI).toFixed(2) : '0.00') + ' | ' +
        (Number.isFinite(entry.maxAQI) ? Number(entry.maxAQI).toFixed(2) : '0.00') + ' | ' +
        (entry.bufferKm == null ? '-' : String(entry.bufferKm))
      );
    });

    const safeProject = (projectFilter === 'all' ? 'all-projects' : projectFilter)
      .toLowerCase()
      .replace(/[^a-z0-9_-]+/g, '_');
    writer.save('all-sketches-analysis-report-' + safeProject + '.pdf');
  }

  exportAnalysisCsvBtn.addEventListener('click', function () {
    if (!lastAnalysisResult) {
      return;
    }

    const rows = analysisRowsToCsv(lastAnalysisResult);
    const csv = rows
      .map(function (row) {
        return row
          .map(function (cell) {
            const text = String(cell).replace(/"/g, '""');
            return '"' + text + '"';
          })
          .join(',');
      })
      .join('\n');

    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = 'sketch-analysis.csv';
    link.click();
    URL.revokeObjectURL(url);
  });

  exportAnalysisPdfBtn.addEventListener('click', function () {
    try {
      exportCurrentSketchAnalysisPdf();
    } catch (err) {
      console.error('Sketch analysis PDF export error:', err);
    }
  });

  function renderAnalysisPopup(analysisResult) {
    lastAnalysisResult = analysisResult;
    setAnalysisPopupCollapsed(false);
    analysisTableBodyEl.innerHTML = '';

    addAnalysisRow('Sketch', 'Geometry type', analysisResult.geometryType);

    if (analysisResult.intersections.length === 0) {
      addAnalysisRow('Intersection', 'Visible selected layers', 'None');
    } else {
      analysisResult.intersections.forEach(function (item) {
        addAnalysisRow('Intersection', item.layerName, item.count + ' feature(s)');
      });
    }

    const hotspotTopBrightness = analysisResult.hotspots.items.reduce(function (maxValue, item) {
      const v = Number(item.brightness);
      return Number.isFinite(v) && v > maxValue ? v : maxValue;
    }, -Infinity);
    const hotspotSeverity = getHotspotSeverity(hotspotTopBrightness);
    addAnalysisRow(
      'Hotspot',
      'Detected hotspots',
      analysisResult.hotspots.count + ' hotspot(s)',
      createBadge(hotspotSeverity.label, hotspotSeverity.bg, hotspotSeverity.fg)
    );

    analysisResult.hotspots.items.slice(0, 3).forEach(function (item, index) {
      const itemSeverity = getHotspotSeverity(item.brightness);
      const value = 'Brightness ' + (item.brightness ?? 'n/a') + ', FRP ' + (item.radiativePower ?? 'n/a') + ', Confidence ' + (item.confidence ?? 'n/a');
      addAnalysisRow(
        'Hotspot #' + (index + 1),
        'Details',
        value,
        createBadge(itemSeverity.label, itemSeverity.bg, itemSeverity.fg)
      );
    });

    addAnalysisRow('Pollution', 'Stations in area', String(analysisResult.pollution.count));
    if (analysisResult.pollution.count > 0) {
      const avgCategory = getAqiCategory(analysisResult.pollution.averageAQI);
      const maxCategory = getAqiCategory(analysisResult.pollution.maxAQI);
      addAnalysisRow(
        'Pollution',
        'Average AQI',
        analysisResult.pollution.averageAQI.toFixed(2),
        createBadge(avgCategory.label, avgCategory.bg, avgCategory.fg)
      );
      addAnalysisRow(
        'Pollution',
        'Max AQI',
        String(analysisResult.pollution.maxAQI),
        createBadge(maxCategory.label, maxCategory.bg, maxCategory.fg)
      );
    }

    if (analysisResult.bufferResults) {
      addAnalysisRow('Buffer', 'Distance', analysisResult.bufferResults.distanceKm + ' km');
      analysisResult.bufferResults.layers.forEach(function (item) {
        addAnalysisRow('Buffer', item.layerName, item.count + ' feature(s)');
      });
    }

    analysisSummaryEl.textContent =
      'Analysis complete for ' + analysisResult.geometryType + '. ' +
      'Intersections: ' + analysisResult.intersections.reduce(function (sum, item) { return sum + item.count; }, 0) +
      ', Hotspots: ' + analysisResult.hotspots.count +
      ', Pollution stations: ' + analysisResult.pollution.count +
      (analysisResult.bufferResults ? ', Buffer layers checked: ' + analysisResult.bufferResults.layers.length : '');

    positionAnalysisPopup();
    analysisPopup.style.display = 'block';
  }

  // listen for complete event and push geometry to Supabase
  sketch.on('create', async function(event) {
    if (event.state === 'complete') {
      const geometry = event.graphic.geometry;
      const wgsGeometry = await toWgs84Geometry(geometry);
      const geojson = arcgisToGeojson(wgsGeometry);

      if (!geojson) {
        console.warn('Unsupported geometry type for saving:', geometry.type);
        return;
      }

      if (!currentUser) {
        await refreshAuthSession();
      }

      if (!currentUser) {
        setAuthStatus('Sign in to save sketches privately. Running analysis only.', '#8a5a00');
        await runSpatialQueries(geometry, { persistBuffer: false });
        return;
      }

      try {
        const insertRow = { geom: geojson };
        if (sketchTableSupportsUserId) {
          insertRow.user_id = currentUser.id;
        }
        if (sketchTableSupportsProjectName) {
          insertRow.project_name = getCurrentProjectNameForSave();
        }

        let { error } = await supabase
          .from('user_sketches')
          .insert([insertRow]);

        if (error && sketchTableSupportsProjectName && isMissingColumnError(error, 'project_name')) {
          console.warn('project_name column not available on user_sketches. Saving without project grouping.');
          sketchTableSupportsProjectName = false;
          setAuthStatus('project_name column missing. Project grouping is temporarily disabled.', '#8a5a00');
          delete insertRow.project_name;
          const retryProject = await supabase
            .from('user_sketches')
            .insert([insertRow]);
          error = retryProject.error;
        }

        if (error && sketchTableSupportsUserId && isMissingColumnError(error, 'user_id')) {
          console.warn('user_id column not available on user_sketches. Saving without user scoping.');
          sketchTableSupportsUserId = false;
          setAuthStatus('user_id column missing. Per-user sketch scoping is disabled.', '#8a5a00');
          delete insertRow.user_id;
          const retryUser = await supabase
            .from('user_sketches')
            .insert([insertRow]);
          error = retryUser.error;
        }

        if (error) {
          console.error('Supabase insert error:', error);
          console.warn('If using row-level security, make sure a policy allows inserts (e.g. create policy "allow insert" on user_sketches for insert with check (true);)');
        } else {
          console.log('Sketch saved to Supabase');
          setAuthStatus('Sketch saved under project "' + getCurrentProjectNameForSave() + '".', '#0f7a30');
          // refresh saved layer and selector from DB source of truth
          await loadSketches();
        }
      } catch (err) {
        console.error('Unexpected error saving sketch:', err);
      }

      // Run spatial queries on the sketched geometry
      await runSpatialQueries(geometry);
    }
  });

  // Spatial Query Functions
  async function runSpatialQueries(geometry, options) {
    const settings = options || {};
    const shouldPersistBuffer = settings.persistBuffer !== false;
    const intersections = await queryIntersections(geometry);
    const hotspots = await queryHotspotDetection(geometry);
    const pollution = await queryPollutionAnalysis(geometry);

    let bufferResults = null;
    if (geometry.type === 'point') {
      bufferResults = await queryBufferAnalysis(geometry);

      if (bufferResults && bufferResults.bufferGeometry) {
        drawSketchAnalysisBuffer(bufferResults.bufferGeometry, geometry);
        if (shouldPersistBuffer) {
          await persistSketchAnalysisBuffer(bufferResults.bufferGeometry, geometry, bufferResults.distanceKm);
        }
      }
    }

    renderAnalysisPopup({
      geometryType: geometry.type,
      intersections: intersections,
      hotspots: hotspots,
      pollution: pollution,
      bufferResults: bufferResults
    });
  }

  // Intersection query - find all features that intersect the sketched geometry
  async function queryIntersections(geometry) {
    const visibleLayers = Object.keys(layerMap).filter(key => layerMap[key].visible);

    const results = await Promise.all(visibleLayers.map(async function(layerId) {
      const layer = layerMap[layerId];
      const queryParams = {
        geometry: geometry,
        spatialRelationship: 'intersects',
        outFields: ['*'],
        returnGeometry: false
      };

      try {
        const queryResult = await layer.queryFeatures(queryParams);
        return {
          layerName: layer.title,
          count: queryResult.features.length
        };
      } catch (err) {
        console.error('Intersection query error for ' + layer.title + ':', err);
        return {
          layerName: layer.title,
          count: 0
        };
      }
    }));

    return results;
  }

  // Hotspot detection - find fire hotspots within or near the sketched geometry
  async function queryHotspotDetection(geometry) {
    const queryParams = {
      geometry: geometry,
      spatialRelationship: 'intersects',
      outFields: ['*'],
      returnGeometry: false
    };

    try {
      const results = await fireHotspotsLayer.queryFeatures(queryParams);
      const items = results.features.map(function(feature) {
        return {
          brightness: feature.attributes.brightness,
          radiativePower: feature.attributes.radiative_power ?? feature.attributes.frp,
          confidence: feature.attributes.confidence
        };
      });

      return {
        count: items.length,
        items: items
      };
    } catch (err) {
      console.error('Hotspot detection error:', err);
      return {
        count: 0,
        items: []
      };
    }
  }

  // Pollution analysis - analyze air quality data within the sketched geometry
  async function queryPollutionAnalysis(geometry) {
    const queryParams = {
      geometry: geometry,
      spatialRelationship: 'intersects',
      outFields: ['*'],
      returnGeometry: false
    };

    try {
      const results = await airQualityLayer.queryFeatures(queryParams);
      const aqiValues = results.features
        .map(function(feature) {
          return Number(feature.attributes.AQI);
        })
        .filter(function(value) {
          return Number.isFinite(value);
        });

      const totalAQI = aqiValues.reduce(function(sum, value) {
        return sum + value;
      }, 0);
      const averageAQI = aqiValues.length > 0 ? totalAQI / aqiValues.length : 0;
      const maxAQI = aqiValues.length > 0 ? Math.max.apply(null, aqiValues) : 0;

      return {
        count: results.features.length,
        averageAQI: averageAQI,
        maxAQI: maxAQI
      };
    } catch (err) {
      console.error('Pollution analysis error:', err);
      return {
        count: 0,
        averageAQI: 0,
        maxAQI: 0
      };
    }
  }

  function getConfiguredBufferDistanceKm() {
    const bufferInputEl = document.getElementById('bufferDistance');
    const configuredKm = bufferInputEl ? Number.parseFloat(bufferInputEl.value) : NaN;
    return Number.isFinite(configuredKm) && configuredKm > 0 ? configuredKm : 5;
  }

  function createAnalysisBufferGeometry(geometry, bufferKm) {
    if (!geometry) {
      return null;
    }

    try {
      if (geometry.spatialReference && (geometry.spatialReference.isGeographic || geometry.spatialReference.isWebMercator)) {
        return geometryEngine.geodesicBuffer(geometry, bufferKm, 'kilometers');
      }
      return geometryEngine.buffer(geometry, bufferKm, 'kilometers');
    } catch (err) {
      // Fallback path for mixed-unit geometries loaded from external sources.
      try {
        return geometryEngine.geodesicBuffer(geometry, bufferKm, 'kilometers');
      } catch (fallbackErr) {
        console.error('Unable to create buffer geometry for analysis:', fallbackErr);
        return null;
      }
    }
  }

  // Buffer analysis - for point geometries, create buffer and find features within
  async function queryBufferAnalysis(geometry) {
    const bufferKm = getConfiguredBufferDistanceKm();
    const bufferGeometry = createAnalysisBufferGeometry(geometry, bufferKm);

    if (!bufferGeometry) {
      return {
        distanceKm: bufferKm,
        bufferGeometry: null,
        layers: []
      };
    }
    
    const visibleLayers = Object.keys(layerMap).filter(key => layerMap[key].visible);

    const layers = await Promise.all(visibleLayers.map(async function(layerId) {
      const layer = layerMap[layerId];
      const queryParams = {
        geometry: bufferGeometry,
        spatialRelationship: 'intersects',
        outFields: ['*'],
        returnGeometry: false
      };

      try {
        const results = await layer.queryFeatures(queryParams);
        return {
          layerName: layer.title,
          count: results.features.length
        };
      } catch (err) {
        console.error('Buffer query error for ' + layer.title + ':', err);
        return {
          layerName: layer.title,
          count: 0
        };
      }
    }));

    return {
      distanceKm: bufferKm,
      bufferGeometry: bufferGeometry,
      layers: layers
    };
  }

  function isMissingColumnError(error, columnName) {
    if (!error || !columnName) {
      return false;
    }

    const text = [error.message, error.details, error.hint]
      .filter(Boolean)
      .join(' ')
      .toLowerCase();

    return text.indexOf('column') !== -1 && text.indexOf(String(columnName).toLowerCase()) !== -1;
  }

  function normalizeProjectName(value) {
    const trimmed = String(value == null ? '' : value).trim();
    return trimmed || defaultProjectName;
  }

  function getRowProjectName(row) {
    if (!sketchTableSupportsProjectName) {
      return defaultProjectName;
    }

    return normalizeProjectName(row && row.project_name);
  }

  function getCurrentProjectNameForSave() {
    return normalizeProjectName(projectNameInputEl ? projectNameInputEl.value : defaultProjectName);
  }

  function getSelectedProjectFilter() {
    if (!projectFilterSelectEl) {
      return 'all';
    }

    return projectFilterSelectEl.value || 'all';
  }

  function applyProjectFilter(rows) {
    const filter = getSelectedProjectFilter();
    if (filter === 'all') {
      return rows.slice();
    }

    return rows.filter(function (row) {
      return getRowProjectName(row) === filter;
    });
  }

  function setAuthStatus(message, color) {
    if (!authStatusEl) {
      return;
    }

    authStatusEl.textContent = message;
    if (color) {
      authStatusEl.style.color = color;
    }
  }

  function updateAuthUiState() {
    const signedIn = !!currentUser;

    if (signedIn) {
      const label = currentUser.email || currentUser.phone || String(currentUser.id || 'user').slice(0, 8);
      setAuthStatus('Signed in as ' + label, '#0f7a30');
      if (authStatusEl) {
        authStatusEl.style.background = '#e6f6ea';
        authStatusEl.style.borderColor = '#b7e3c4';
      }
    } else {
      setAuthStatus('Not signed in. Use Sign In below to load private sketches.', '#264d86');
      if (authStatusEl) {
        authStatusEl.style.background = '#f4f7fb';
        authStatusEl.style.borderColor = '#c8d3e6';
      }
    }

    if (authSignInBtnEl) {
      authSignInBtnEl.style.display = signedIn ? 'none' : 'inline-flex';
    }

    if (authSignUpBtnEl) {
      authSignUpBtnEl.style.display = signedIn ? 'none' : 'inline-flex';
    }

    if (authSignOutBtnEl) {
      authSignOutBtnEl.style.display = signedIn ? 'inline-flex' : 'none';
    }

    if (projectNameInputEl) {
      projectNameInputEl.disabled = !signedIn;
      if (!signedIn) {
        projectNameInputEl.value = defaultProjectName;
      }
    }

    if (projectFilterSelectEl) {
      projectFilterSelectEl.disabled = !signedIn;
    }
  }

  function promptForCredentials(actionLabel) {
    const emailInput = window.prompt('Email for ' + actionLabel + ':', '');
    if (emailInput == null) {
      return null;
    }

    const passwordInput = window.prompt('Password for ' + actionLabel + ':', '');
    if (passwordInput == null) {
      return null;
    }

    return {
      email: String(emailInput || '').trim(),
      password: String(passwordInput || '')
    };
  }

  function refreshProjectFilterSelect() {
    if (!projectFilterSelectEl) {
      return;
    }

    const previous = projectFilterSelectEl.value || 'all';
    projectFilterSelectEl.innerHTML = '';

    const allOption = document.createElement('option');
    allOption.value = 'all';
    allOption.textContent = 'All projects';
    projectFilterSelectEl.appendChild(allOption);

    const uniqueProjects = [];
    const seen = new Set();
    allSavedSketchRows.forEach(function (row) {
      const projectName = getRowProjectName(row);
      if (!seen.has(projectName)) {
        seen.add(projectName);
        uniqueProjects.push(projectName);
      }
    });

    uniqueProjects.sort(function (a, b) {
      return a.localeCompare(b);
    });

    uniqueProjects.forEach(function (projectName) {
      const option = document.createElement('option');
      option.value = projectName;
      option.textContent = projectName;
      projectFilterSelectEl.appendChild(option);
    });

    if (previous !== 'all' && uniqueProjects.indexOf(previous) === -1) {
      projectFilterSelectEl.value = 'all';
    } else {
      projectFilterSelectEl.value = previous;
    }
  }

  async function refreshAuthSession() {
    try {
      const { data, error } = await supabase.auth.getUser();
      if (error) {
        console.warn('Auth session check warning:', error.message || error);
        currentUser = null;
      } else {
        currentUser = data && data.user ? data.user : null;
      }
    } catch (err) {
      console.warn('Unable to refresh auth session:', err);
      currentUser = null;
    }

    updateAuthUiState();
  }

  async function signInFromInputs() {
    const credentials = promptForCredentials('sign in');
    if (!credentials) {
      return;
    }

    if (!credentials.email || !credentials.password) {
      alert('Enter email and password.');
      return;
    }

    const { error } = await supabase.auth.signInWithPassword({ email: credentials.email, password: credentials.password });
    if (error) {
      alert('Sign in failed: ' + (error.message || 'unknown error'));
      return;
    }

    await refreshAuthSession();
    await loadSketches();
  }

  async function signUpFromInputs() {
    const credentials = promptForCredentials('sign up');
    if (!credentials) {
      return;
    }

    if (!credentials.email || !credentials.password) {
      alert('Enter email and password.');
      return;
    }

    const { error } = await supabase.auth.signUp({ email: credentials.email, password: credentials.password });
    if (error) {
      alert('Sign up failed: ' + (error.message || 'unknown error'));
      return;
    }

    alert('Sign up succeeded. If email confirmation is enabled, verify your email before signing in.');
    await refreshAuthSession();
    await loadSketches();
  }

  async function signOutCurrentUser() {
    const { error } = await supabase.auth.signOut();
    if (error) {
      alert('Sign out failed: ' + (error.message || 'unknown error'));
      return;
    }

    currentUser = null;
    updateAuthUiState();
    await loadSketches();
  }

  // Load sketches from Supabase and display on map
  async function loadSketches() {
    try {
      if (!currentUser) {
        savedSketchesLayer.removeAll();
        savedSketchRecords = [];
        allSavedSketchRows = [];
        lastVisibleSavedCount = 0;
        refreshProjectFilterSelect();
        refreshSavedSketchSelect();
        return;
      }

      let query = supabase
        .from('user_sketches')
        .select('*')
        .order('created_at', { ascending: true });

      if (sketchTableSupportsUserId) {
        query = query.eq('user_id', currentUser.id);
      }

      let { data, error } = await query;

      if (error && sketchTableSupportsUserId && isMissingColumnError(error, 'user_id')) {
        console.warn('user_id column not available on user_sketches. Falling back to non-user-scoped query.');
        sketchTableSupportsUserId = false;
        setAuthStatus('user_id column missing. Showing unscoped sketches.', '#8a5a00');
        const retry = await supabase
          .from('user_sketches')
          .select('*')
          .order('created_at', { ascending: true });
        data = retry.data;
        error = retry.error;
      }

      if (error) {
        console.error('Error fetching sketches:', error);
        return;
      }
      if (!data) {
        console.warn('Fetch returned no data (network issue?)');
        return;
      }

      allSavedSketchRows = data.slice();
      refreshProjectFilterSelect();
      const filteredRows = applyProjectFilter(allSavedSketchRows);

      savedSketchesLayer.removeAll();
      savedSketchRecords = [];
      let renderedCount = 0;

      filteredRows.forEach(function (feature) {
        const geometry = geojsonToArcgis(feature.geom);
        if (!geometry) {
          console.warn('Skipped sketch row due to unsupported geom format:', feature.id, feature.geom);
          return;
        }

        const graphic = new Graphic({
          geometry: geometry,
          symbol: getSavedSketchSymbol(geometry.type),
          attributes: {
            id: feature.id,
            created_at: feature.created_at,
            geometry_type: geometry.type,
            project_name: getRowProjectName(feature)
          }
        });
        savedSketchesLayer.add(graphic);
        savedSketchRecords.push({
          id: feature.id,
          created_at: feature.created_at,
          geometry_type: geometry.type,
          project_name: getRowProjectName(feature),
          user_id: feature.user_id,
          graphic: graphic
        });
        renderedCount += 1;
      });

      lastVisibleSavedCount = filteredRows.length;
      refreshSavedSketchSelect();

      console.log('Loaded', renderedCount, 'of', filteredRows.length, 'sketches from Supabase (project filter:', getSelectedProjectFilter() + ')');
      if (data.length === 0) {
        console.warn('No saved sketches found for this user.');
      }
      if (filteredRows.length > 0 && renderedCount === 0) {
        console.warn('Rows exist but none could be rendered. Check geom format or RLS select policy.');
      }
    } catch (err) {
      console.error('Unexpected error loading sketches:', err);
    }
  }

  function refreshSavedSketchSelect() {
    if (!savedSketchSelectEl) {
      return;
    }

    const previousValue = savedSketchSelectEl.value;
    savedSketchSelectEl.innerHTML = '';

    const placeholder = document.createElement('option');
    placeholder.value = '';
    if (!currentUser) {
      placeholder.textContent = 'Sign in to load your sketches';
    } else if (savedSketchRecords.length > 0) {
      placeholder.textContent = 'Choose a saved sketch...';
    } else if (allSavedSketchRows.length > 0 && getSelectedProjectFilter() !== 'all') {
      placeholder.textContent = 'No sketches in selected project';
    } else if (lastVisibleSavedCount === 0) {
      placeholder.textContent = 'No saved sketches';
    } else {
      placeholder.textContent = 'No visible sketches';
    }
    savedSketchSelectEl.appendChild(placeholder);

    savedSketchRecords.forEach(function (record, index) {
      const option = document.createElement('option');
      option.value = String(index);

      let createdText = 'no timestamp';
      if (record.created_at) {
        const dt = new Date(record.created_at);
        createdText = Number.isNaN(dt.getTime()) ? record.created_at : dt.toLocaleString();
      }

      option.textContent = (index + 1) + '. [' + normalizeProjectName(record.project_name) + '] ' + (record.geometry_type || 'geometry') + ' - ' + createdText;
      savedSketchSelectEl.appendChild(option);
    });

    if (previousValue !== '' && savedSketchRecords[Number(previousValue)]) {
      savedSketchSelectEl.value = previousValue;
    }
  }

  function getSelectedSavedSketchRecord() {
    const selectedIndex = Number(savedSketchSelectEl ? savedSketchSelectEl.value : NaN);
    if (!Number.isInteger(selectedIndex) || !savedSketchRecords[selectedIndex]) {
      return null;
    }

    return savedSketchRecords[selectedIndex];
  }

  function getSavedSketchRecordIndex(record) {
    if (!record) {
      return -1;
    }

    return savedSketchRecords.findIndex(function (item) {
      return String(item.id) === String(record.id);
    });
  }

  function findSavedSketchRecordByGraphic(graphic) {
    if (!graphic) {
      return null;
    }

    const attrId = graphic.attributes && graphic.attributes.id;
    if (attrId != null) {
      const byId = savedSketchRecords.find(function (record) {
        return String(record.id) === String(attrId);
      });
      if (byId) {
        return byId;
      }
    }

    const byRef = savedSketchRecords.find(function (record) {
      return record.graphic === graphic;
    });

    return byRef || null;
  }

  async function analyzeSavedSketchRecord(record, options) {
    if (!record || !record.graphic || !record.graphic.geometry) {
      return;
    }

    const settings = options || {};

    if (savedSketchSelectEl) {
      const selectedIndex = getSavedSketchRecordIndex(record);
      if (selectedIndex >= 0) {
        savedSketchSelectEl.value = String(selectedIndex);
      }
    }

    if (settings.zoom !== false) {
      try {
        await view.goTo(record.graphic.geometry);
      } catch (err) {
        console.warn('Unable to zoom to selected sketch:', err);
      }
    }

    await runSpatialQueries(record.graphic.geometry, { persistBuffer: false });
  }

  async function goToAndAnalyzeSelectedSketch() {
    const selected = getSelectedSavedSketchRecord();
    if (!selected) {
      alert('Select a saved sketch first.');
      return;
    }

    await analyzeSavedSketchRecord(selected, { zoom: true });
  }

  // call loadSketches after view is ready
  view.when(async function() {
    console.log('MapView is ready, spatial reference:', view.spatialReference, 'wkid:', view.spatialReference.wkid);
    // add collapsible utility controls in map UI
    const controls = document.createElement('div');
    controls.id = 'savedSketchToolsWidget';
    controls.className = 'esri-widget esri-component';
    controls.style.cssText = 'display:flex;flex-direction:column;gap:6px;padding:8px;background:rgba(255,255,255,0.96);border:1px solid #dde4f0;border-radius:8px;box-shadow:0 4px 12px rgba(0,0,0,0.12);max-width:260px;min-width:240px;';

    const toggleBtn = document.createElement('button');
    toggleBtn.type = 'button';
    toggleBtn.textContent = 'Saved Sketch Tools ▶';
    toggleBtn.style.cssText = 'display:block;width:100%;margin:0;padding:0.45rem 0.6rem;background:#17365d;color:#fff;border:0;border-radius:6px;cursor:pointer;font-size:0.84rem;font-weight:700;text-align:left;';

    const controlsBody = document.createElement('div');
    controlsBody.style.cssText = 'display:none;flex-direction:column;gap:6px;';

    let controlsExpanded = false;
    toggleBtn.addEventListener('click', function (event) {
      event.preventDefault();
      event.stopPropagation();
      controlsExpanded = !controlsExpanded;
      controlsBody.style.display = controlsExpanded ? 'flex' : 'none';
      toggleBtn.textContent = controlsExpanded ? 'Saved Sketch Tools ▼' : 'Saved Sketch Tools ▶';
    });

    function makeBtn(label, onClick) {
      const b = document.createElement('button');
      b.type = 'button';
      b.textContent = label;
      b.style.cssText = 'display:block;width:100%;margin:0;padding:0.4rem 0.8rem;background:#17365d;color:white;border:none;border-radius:4px;cursor:pointer;font-size:0.83rem;';
      b.addEventListener('click', function (event) {
        event.preventDefault();
        event.stopPropagation();
        Promise.resolve(onClick()).catch(function (err) {
          console.error('Control action error:', err);
        });
      });
      return b;
    }

    function makeInlineBtn(label, onClick, background) {
      const b = document.createElement('button');
      b.type = 'button';
      b.textContent = label;
      b.style.cssText = 'display:inline-flex;align-items:center;justify-content:center;padding:0.28rem 0.45rem;background:' + (background || '#17365d') + ';color:white;border:none;border-radius:4px;cursor:pointer;font-size:0.74rem;line-height:1.1;';
      b.addEventListener('click', function (event) {
        event.preventDefault();
        event.stopPropagation();
        Promise.resolve(onClick()).catch(function (err) {
          console.error('Inline control action error:', err);
        });
      });
      return b;
    }

    const authLabel = document.createElement('label');
    authLabel.textContent = 'Account';
    authLabel.style.cssText = 'display:block;font-size:0.78rem;color:#264d86;font-weight:700;margin-top:2px;';
    controlsBody.appendChild(authLabel);

    authStatusEl = document.createElement('div');
    authStatusEl.style.cssText = 'font-size:0.74rem;color:#264d86;background:#f4f7fb;border:1px solid #c8d3e6;border-radius:4px;padding:0.32rem 0.45rem;line-height:1.25;';
    controlsBody.appendChild(authStatusEl);

    const authButtonsRow = document.createElement('div');
    authButtonsRow.style.cssText = 'display:flex;gap:0.35rem;flex-wrap:wrap;align-items:center;';
    authSignInBtnEl = makeInlineBtn('Sign In', function () {
      return signInFromInputs();
    }, '#17365d');
    authButtonsRow.appendChild(authSignInBtnEl);

    authSignUpBtnEl = makeInlineBtn('Create Account', function () {
      return signUpFromInputs();
    }, '#264d86');
    authButtonsRow.appendChild(authSignUpBtnEl);

    authSignOutBtnEl = makeInlineBtn('Sign Out', function () {
      return signOutCurrentUser();
    }, '#7a1a1a');
    authButtonsRow.appendChild(authSignOutBtnEl);
    controlsBody.appendChild(authButtonsRow);

    const projectSaveLabel = document.createElement('label');
    projectSaveLabel.textContent = 'Project for new sketches';
    projectSaveLabel.style.cssText = 'display:block;font-size:0.78rem;color:#264d86;font-weight:600;margin-top:4px;';
    controlsBody.appendChild(projectSaveLabel);

    projectNameInputEl = document.createElement('input');
    projectNameInputEl.type = 'text';
    projectNameInputEl.value = defaultProjectName;
    projectNameInputEl.placeholder = defaultProjectName;
    projectNameInputEl.style.cssText = 'display:block;width:100%;margin:0;padding:0.35rem;border:1px solid #c8d3e6;border-radius:4px;background:#fff;color:#17365d;font-size:0.8rem;';
    projectNameInputEl.addEventListener('change', function () {
      projectNameInputEl.value = normalizeProjectName(projectNameInputEl.value);
    });
    controlsBody.appendChild(projectNameInputEl);

    const projectFilterLabel = document.createElement('label');
    projectFilterLabel.textContent = 'Project filter';
    projectFilterLabel.style.cssText = 'display:block;font-size:0.78rem;color:#264d86;font-weight:600;margin-top:2px;';
    controlsBody.appendChild(projectFilterLabel);

    projectFilterSelectEl = document.createElement('select');
    projectFilterSelectEl.style.cssText = 'display:block;width:100%;margin:0;padding:0.35rem;border:1px solid #c8d3e6;border-radius:4px;background:#fff;color:#17365d;font-size:0.82rem;';
    projectFilterSelectEl.addEventListener('change', function () {
      Promise.resolve(loadSketches()).catch(function (err) {
        console.error('Project filter change error:', err);
      });
    });
    controlsBody.appendChild(projectFilterSelectEl);

    refreshProjectFilterSelect();
    updateAuthUiState();

    const selectorLabel = document.createElement('label');
    selectorLabel.textContent = 'Saved sketches';
    selectorLabel.style.cssText = 'display:block;font-size:0.78rem;color:#264d86;font-weight:600;margin-top:2px;';

    savedSketchSelectEl = document.createElement('select');
    savedSketchSelectEl.style.cssText = 'display:block;width:100%;margin:0;padding:0.35rem;border:1px solid #c8d3e6;border-radius:4px;background:#fff;color:#17365d;font-size:0.82rem;';
    refreshSavedSketchSelect();
    savedSketchSelectEl.addEventListener('change', function () {
      if (!savedSketchSelectEl.value) {
        return;
      }

      Promise.resolve(goToAndAnalyzeSelectedSketch()).catch(function (err) {
        console.error('Selected-sketch analysis error:', err);
      });
    });

    controlsBody.appendChild(selectorLabel);
    controlsBody.appendChild(savedSketchSelectEl);

    controlsBody.appendChild(makeBtn('All sketches analysis', function () {
      openAllSketchesAnalysisPopup();
    }));

    controlsBody.appendChild(makeBtn('Go to + analyze selected', function () {
      return goToAndAnalyzeSelectedSketch();
    }));

    controlsBody.appendChild(makeBtn('Clear current sketches', function () {
      sketchGraphicsLayer.removeAll();
      if (typeof sketch.cancel === 'function') {
        sketch.cancel();
      }
    }));

    controlsBody.appendChild(makeBtn('Clear analysis buffers', function () {
      sketchBufferGraphicsLayer.removeAll();
      setPersistedSketchBuffers([]);
    }));

    controlsBody.appendChild(makeBtn('Delete all saved', async function () {
      if (!currentUser) {
        alert('Sign in first to delete your saved sketches.');
        return;
      }

      if (!confirm('Remove all stored sketches from Supabase?')) {
        return;
      }

      const idsToDelete = savedSketchRecords
        .map(function (record) { return record.id; })
        .filter(function (id) { return !!id; });

      if (idsToDelete.length === 0) {
        alert('No visible saved sketches to delete for current project filter.');
        return;
      }

      let deleteQuery = supabase
        .from('user_sketches')
        .delete()
        .in('id', idsToDelete);

      if (sketchTableSupportsUserId) {
        deleteQuery = deleteQuery.eq('user_id', currentUser.id);
      }

      const activeProjectFilter = getSelectedProjectFilter();
      if (sketchTableSupportsProjectName && activeProjectFilter !== 'all') {
        deleteQuery = deleteQuery.eq('project_name', activeProjectFilter);
      }

      let { data: deletedRows, error } = await deleteQuery.select('id');

      if (error && (isMissingColumnError(error, 'user_id') || isMissingColumnError(error, 'project_name'))) {
        if (isMissingColumnError(error, 'user_id')) {
          sketchTableSupportsUserId = false;
        }
        if (isMissingColumnError(error, 'project_name')) {
          sketchTableSupportsProjectName = false;
        }

        const retry = await supabase
          .from('user_sketches')
          .delete()
          .in('id', idsToDelete)
          .select('id');
        deletedRows = retry.data;
        error = retry.error;
      }

      if (error) {
        console.error('delete error', error);
        alert('Delete failed. Check RLS DELETE policy on user_sketches.');
        return;
      }

      const deletedCount = Array.isArray(deletedRows) ? deletedRows.length : 0;
      if (deletedCount === 0) {
        alert('No rows were deleted. This usually means RLS DELETE policy is missing for anon.');
      } else {
        alert('Deleted ' + deletedCount + ' saved sketch(es).');
      }

      sketchBufferGraphicsLayer.removeAll();
      setPersistedSketchBuffers([]);
      await loadSketches();
    }));

    controlsBody.appendChild(makeBtn('Export saved GeoJSON', async function () {
      if (!currentUser) {
        alert('Sign in first to export your saved sketches.');
        return;
      }

      const projectFilter = getSelectedProjectFilter();
      const exportColumns = sketchTableSupportsProjectName ? 'id, created_at, geom, project_name' : 'id, created_at, geom';

      let exportQuery = supabase
        .from('user_sketches')
        .select(exportColumns)
        .order('created_at', { ascending: true });

      if (sketchTableSupportsUserId) {
        exportQuery = exportQuery.eq('user_id', currentUser.id);
      }

      let { data, error } = await exportQuery;

      if (error && sketchTableSupportsUserId && isMissingColumnError(error, 'user_id')) {
        sketchTableSupportsUserId = false;
        const retryNoUser = await supabase
          .from('user_sketches')
          .select(exportColumns)
          .order('created_at', { ascending: true });
        data = retryNoUser.data;
        error = retryNoUser.error;
      }

      if (error && sketchTableSupportsProjectName && isMissingColumnError(error, 'project_name')) {
        sketchTableSupportsProjectName = false;
        const retryNoProject = await supabase
          .from('user_sketches')
          .select('id, created_at, geom')
          .order('created_at', { ascending: true });
        data = retryNoProject.data;
        error = retryNoProject.error;
      }

      if (error) {
        console.error('Export fetch error:', error);
        return;
      }

      const scopedRows = (data || []).filter(function (row) {
        return projectFilter === 'all' || getRowProjectName(row) === projectFilter;
      });

      const features = scopedRows
        .map(function (row) {
          const geometry = normalizeStoredGeometry(row.geom);
          if (!geometry) return null;
          return {
            type: 'Feature',
            geometry: geometry,
            properties: {
              id: row.id,
              created_at: row.created_at,
              project_name: getRowProjectName(row)
            }
          };
        })
        .filter(Boolean);

      console.log('exporting', features.length, 'saved sketches for project filter', projectFilter);
      if (scopedRows.length > 0 && features.length === 0) {
        console.warn('Export found rows but could not parse geom. Confirm geom is valid GeoJSON/jsonb or supported WKT.');
      }

      const blob = new Blob([JSON.stringify({ type: 'FeatureCollection', features: features }, null, 2)], { type: 'application/json' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      const safeProject = (projectFilter === 'all' ? 'all-projects' : projectFilter)
        .toLowerCase()
        .replace(/[^a-z0-9_-]+/g, '_');
      a.download = 'sketches-' + safeProject + '.geojson';
      a.click();
      URL.revokeObjectURL(url);
    }));

    controls.appendChild(toggleBtn);
    controls.appendChild(controlsBody);
    view.ui.add(controls, 'bottom-left');

    // map.spatialReference may be undefined; use view.spatialReference when needed
    restorePersistedSketchBuffers();

    if (!authStateSubscription && supabase && supabase.auth && typeof supabase.auth.onAuthStateChange === 'function') {
      const registration = supabase.auth.onAuthStateChange(function (_event, session) {
        currentUser = session && session.user ? session.user : null;
        updateAuthUiState();
        Promise.resolve(loadSketches()).catch(function (err) {
          console.error('Auth-state sketch reload error:', err);
        });
      });

      authStateSubscription = registration && registration.data ? registration.data.subscription : null;
    }

    await refreshAuthSession();
    await loadSketches();
    schedulePanelToggleRepositionPasses();
  });

  const legend = new Legend({ view: view });
  view.ui.add(new Expand({ view: view, content: legend, expanded: true }), 'top-right');


  // Layer mapping for checkboxes
  const layerMap = {
    'air-quality': airQualityLayer,
    'fire-hotspots': fireHotspotsLayer,
    'industrial-emissions': industrialEmissionsLayer,
    'flood-risk': floodRiskLayer,
    'population-density': populationDensityLayer,
    'hospitals': hospitalsLayer,
    'esg-compliance': esgComplianceLayer,
    'environmental-violations': environmentalViolationsLayer
  };

  function handleCheckboxChange(event) {
    const layerId = event.target.value;
    const isChecked = event.target.checked;
    
    if (layerMap[layerId]) {
      layerMap[layerId].visible = isChecked;
    }
  }

  const layerCheckboxes = document.querySelectorAll('.layer-checkbox');
  layerCheckboxes.forEach(function (checkbox) {
    checkbox.addEventListener('change', handleCheckboxChange);
  });

  const panelEl = document.getElementById('panel');
  const panelToggleBtn = document.getElementById('panelToggleBtn');
  const panelCollapsedStorageKey = 'esg_panel_collapsed_v1';

  function getPanelToggleIconMarkup(collapsed) {
    if (collapsed) {
      return '<svg viewBox="0 0 24 24" aria-hidden="true" focusable="false"><rect x="3" y="4" width="6" height="16" rx="1.4" fill="currentColor" opacity="0.28"></rect><path d="M13 8l4 4-4 4" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"></path></svg>';
    }

    return '<svg viewBox="0 0 24 24" aria-hidden="true" focusable="false"><rect x="3" y="4" width="6" height="16" rx="1.4" fill="currentColor" opacity="0.28"></rect><path d="M17 8l-4 4 4 4" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"></path></svg>';
  }

  function getTopLeftWidgetsBottom() {
    const topLeft = document.querySelector('.esri-ui-top-left');
    if (!topLeft) {
      return 12;
    }

    let maxBottom = 12;
    const components = topLeft.querySelectorAll('.esri-component');
    components.forEach(function (node) {
      const rect = node.getBoundingClientRect();
      if (rect.width <= 0 || rect.height <= 0) {
        return;
      }

      const computed = window.getComputedStyle(node);
      if (computed.display === 'none' || computed.visibility === 'hidden') {
        return;
      }

      maxBottom = Math.max(maxBottom, Math.ceil(rect.bottom));
    });

    return maxBottom;
  }

  function getSavedSketchToolsRect() {
    const toolsEl = document.getElementById('savedSketchToolsWidget');
    if (!toolsEl) {
      return null;
    }

    const rect = toolsEl.getBoundingClientRect();
    if (rect.width <= 0 || rect.height <= 0) {
      return null;
    }

    return rect;
  }

  function adjustTogglePositionToAvoidSavedTools(left, top) {
    if (!panelToggleBtn) {
      return { left: left, top: top };
    }

    const toolsRect = getSavedSketchToolsRect();
    if (!toolsRect) {
      return { left: left, top: top };
    }

    const buttonWidth = panelToggleBtn.offsetWidth || 34;
    const buttonHeight = panelToggleBtn.offsetHeight || 34;
    const toggleRect = {
      left: left,
      top: top,
      right: left + buttonWidth,
      bottom: top + buttonHeight
    };

    const overlaps = !(
      toggleRect.right <= toolsRect.left ||
      toggleRect.left >= toolsRect.right ||
      toggleRect.bottom <= toolsRect.top ||
      toggleRect.top >= toolsRect.bottom
    );

    if (!overlaps) {
      return { left: left, top: top };
    }

    const aboveTop = Math.round(toolsRect.top - buttonHeight - 8);
    if (aboveTop >= 8) {
      return { left: left, top: aboveTop };
    }

    const maxLeft = Math.max(8, window.innerWidth - buttonWidth - 8);
    const sideLeft = Math.min(maxLeft, Math.round(toolsRect.right + 8));
    const clampedTop = Math.min(Math.max(8, top), Math.max(8, window.innerHeight - buttonHeight - 8));
    return { left: sideLeft, top: clampedTop };
  }

  function positionPanelToggleButton() {
    if (!panelToggleBtn) {
      return;
    }

    const isCollapsed = document.body.classList.contains('panel-collapsed');
    panelToggleBtn.style.transform = 'none';
    panelToggleBtn.style.right = 'auto';

    if (!isCollapsed && panelEl) {
      const panelRect = panelEl.getBoundingClientRect();
      const buttonWidth = panelToggleBtn.offsetWidth || 34;
      // Use clientWidth so the toggle stays left of the panel's vertical scrollbar.
      const contentRight = panelRect.left + panelEl.clientWidth;
      const left = Math.max(8, Math.round(contentRight - buttonWidth - 10));
      const top = Math.max(8, Math.round(panelRect.top + 10));

      panelToggleBtn.style.left = left + 'px';
      panelToggleBtn.style.top = top + 'px';
      return;
    }

    const widgetsBottom = getTopLeftWidgetsBottom();
    const baseLeft = 12;
    const baseTop = Math.max(12, widgetsBottom + 8);
    const adjusted = adjustTogglePositionToAvoidSavedTools(baseLeft, baseTop);
    panelToggleBtn.style.left = adjusted.left + 'px';
    panelToggleBtn.style.top = adjusted.top + 'px';
  }

  function schedulePanelToggleRepositionPasses() {
    [80, 220, 500, 900, 1500].forEach(function (delay) {
      setTimeout(function () {
        positionPanelToggleButton();
      }, delay);
    });
  }

  function applyPanelCollapsedState(collapsed) {
    document.body.classList.toggle('panel-collapsed', collapsed);

    if (panelToggleBtn) {
      panelToggleBtn.innerHTML = getPanelToggleIconMarkup(collapsed);
      panelToggleBtn.setAttribute('aria-label', collapsed ? 'Expand ESG Explorer' : 'Collapse ESG Explorer');
      panelToggleBtn.setAttribute('title', collapsed ? 'Expand ESG Explorer' : 'Collapse ESG Explorer');
    }

    try {
      localStorage.setItem(panelCollapsedStorageKey, collapsed ? '1' : '0');
    } catch (_err) {
      // Ignore storage failures.
    }

    positionPanelToggleButton();
    schedulePanelToggleRepositionPasses();

    window.dispatchEvent(new Event('resize'));
    setTimeout(function () {
      positionPanelToggleButton();
      window.dispatchEvent(new Event('resize'));
    }, 300);
  }

  if (panelEl && panelToggleBtn) {
    let initiallyCollapsed = false;
    try {
      initiallyCollapsed = localStorage.getItem(panelCollapsedStorageKey) === '1';
    } catch (_err) {
      initiallyCollapsed = false;
    }

    applyPanelCollapsedState(initiallyCollapsed);
    schedulePanelToggleRepositionPasses();

    panelToggleBtn.addEventListener('click', function () {
      const isCollapsed = document.body.classList.contains('panel-collapsed');
      applyPanelCollapsedState(!isCollapsed);
    });

    window.addEventListener('resize', positionPanelToggleButton);
    window.addEventListener('orientationchange', positionPanelToggleButton);
  }

  // ====== COLLAPSIBLE SECTIONS ======
  const collapsibleHeaders = document.querySelectorAll('.collapsible-header');
  collapsibleHeaders.forEach(function (header) {
    header.addEventListener('click', function () {
      const targetId = this.getAttribute('data-target');
      const content = document.getElementById(targetId);
      const indicator = this.querySelector('.toggle-indicator');

      if (content.classList.contains('collapsed')) {
        // Expand
        content.classList.remove('collapsed');
        indicator.textContent = '▼';
      } else {
        // Collapse
        content.classList.add('collapsed');
        indicator.textContent = '▶';
      }
    });
  });

  // Buffer analysis functionality
  let bufferModeActive = false;
  const bufferDistanceInput = document.getElementById('bufferDistance');
  const activateBufferBtn = document.getElementById('activateBuffer');
  const clearBufferBtn = document.getElementById('clearBuffer');
  const bufferStatus = document.getElementById('bufferStatus');
  const bufferPopup = document.getElementById('bufferPopup');
  const togglePopupBtn = document.getElementById('togglePopup');
  const featuresTableBody = document.getElementById('featuresTableBody');
  const featureDetails = document.getElementById('featureDetails');
  const detailsTitle = document.getElementById('detailsTitle');
  const detailsTableHead = document.getElementById('detailsTableHead');
  const detailsTableBody = document.getElementById('detailsTableBody');
  const closeDetailsBtn = document.getElementById('closeDetails');

  activateBufferBtn.addEventListener('click', function () {
    bufferModeActive = !bufferModeActive;
    if (bufferModeActive) {
      activateBufferBtn.textContent = 'Deactivate Buffer Mode';
      activateBufferBtn.style.background = '#d7191c';
      bufferStatus.textContent = 'Click on map to create buffer';
      clearBufferBtn.style.display = 'block';
      view.cursor = 'crosshair';
    } else {
      activateBufferBtn.textContent = 'Activate Buffer Mode';
      activateBufferBtn.style.background = '#17365d';
      bufferStatus.textContent = '';
      clearBufferBtn.style.display = 'none';
      view.cursor = 'auto';
    }
  });

  clearBufferBtn.addEventListener('click', function () {
    bufferGraphicsLayer.removeAll();
    bufferPopup.classList.add('hidden');
    featuresTableBody.innerHTML = '';
    featureDetails.classList.add('hidden');
    bufferStatus.textContent = 'Buffer cleared. Click on map to create new buffer';
  });

  view.on('click', function (event) {
    if (!bufferModeActive) return;

    const clickPoint = event.mapPoint;
    const bufferDistance = parseFloat(bufferDistanceInput.value) || 5;
    const bufferKm = bufferDistance;

    // Clear previous buffer and results
    bufferGraphicsLayer.removeAll();
    featuresTableBody.innerHTML = '';
    featureDetails.classList.add('hidden');

    // Create a safe buffer that supports geographic and projected points.
    const bufferGeometry = createAnalysisBufferGeometry(clickPoint, bufferKm);

    if (!bufferGeometry) {
      bufferStatus.textContent = 'Unable to create buffer for this location/spatial reference.';
      return;
    }

    // Add buffer to graphics layer
    const bufferGraphic = new Graphic({
      geometry: bufferGeometry,
      symbol: {
        type: 'simple-fill',
        color: 'rgba(0, 150, 200, 0.1)',
        outline: {
          color: 'rgba(0, 150, 200, 0.8)',
          width: 2
        }
      }
    });
    bufferGraphicsLayer.add(bufferGraphic);

    // Add center point
    const centerGraphic = new Graphic({
      geometry: clickPoint,
      symbol: {
        type: 'simple-marker',
        color: '#d7191c',
        size: 8,
        outline: { color: 'white', width: 2 }
      }
    });
    bufferGraphicsLayer.add(centerGraphic);

    // Query features in buffer for each visible layer
    const resultsData = [];
    const visibleLayers = Object.keys(layerMap).filter(key => layerMap[key].visible);

    if (visibleLayers.length === 0) {
      bufferStatus.textContent = 'No layers selected. Please select at least one layer.';
      return;
    }

    let queriesCompleted = 0;

    visibleLayers.forEach(function (layerId) {
      const layer = layerMap[layerId];
      const queryParams = {
        geometry: bufferGeometry,
        spatialRelationship: 'intersects',
        outFields: ['*'],
        returnGeometry: false
      };

      layer.queryFeatures(queryParams).then(function (results) {
        const featureCount = results.features.length;
        resultsData.push({
          layId: layerId,
          layerName: layer.title,
          count: featureCount,
          features: results.features
        });

        queriesCompleted++;
        if (queriesCompleted === visibleLayers.length) {
          displayResults(resultsData);
        }
      }).catch(function (err) {
        console.error('Query error for ' + layer.title + ':', err);
        queriesCompleted++;
        if (queriesCompleted === visibleLayers.length) {
          displayResults(resultsData);
        }
      });
    });
  });

  function displayResults(resultsData) {
    featuresTableBody.innerHTML = '';
    let totalFeatures = 0;

    resultsData.forEach(function (result) {
      totalFeatures += result.count;
      const row = document.createElement('tr');
      
      const layerCell = document.createElement('td');
      layerCell.textContent = result.layerName;
      row.appendChild(layerCell);
      
      const countCell = document.createElement('td');
      countCell.textContent = result.count;
      row.appendChild(countCell);
      
      const detailsCell = document.createElement('td');
      if (result.count > 0) {
        const detailsBtn = document.createElement('button');
        detailsBtn.textContent = 'View Details';
        detailsBtn.className = 'details-btn';
        detailsBtn.addEventListener('click', function() {
          showFeatureDetails(result.features, result.layerName);
        });
        detailsCell.appendChild(detailsBtn);
      } else {
        detailsCell.textContent = 'No features';
      }
      row.appendChild(detailsCell);
      
      featuresTableBody.appendChild(row);
    });

    if (totalFeatures > 0) {
      bufferPopup.classList.remove('hidden');
    }
    bufferStatus.textContent = 'Total: ' + totalFeatures + ' features found in ' + parseFloat(bufferDistanceInput.value) + ' km buffer';
  }

  // ====== MEASUREMENT TOOLS ======
  let distanceModeActive = false;
  let areaModeActive = false;
  let distancePoints = [];
  let areaPoints = [];
  let hoveredSketchGraphic = null;
  let hoveredSketchOriginalSymbol = null;
  let hoverHitTestRequestId = 0;

  function isMapInteractionBusy() {
    return bufferModeActive || distanceModeActive || areaModeActive ||
      (sketch && sketch.viewModel && sketch.viewModel.state === 'active');
  }

  function cloneGraphicSymbol(symbol) {
    if (!symbol) {
      return null;
    }

    if (typeof symbol.clone === 'function') {
      try {
        return symbol.clone();
      } catch (err) {
        // Fall through to JSON clone.
      }
    }

    try {
      return JSON.parse(JSON.stringify(symbol));
    } catch (err) {
      return symbol;
    }
  }

  function getHoverSymbolForGraphic(graphic) {
    if (!graphic || !graphic.geometry) {
      return null;
    }

    if (graphic.geometry.type === 'point') {
      const baseSize = Number(graphic.symbol && graphic.symbol.size);
      return {
        type: 'simple-marker',
        color: [255, 140, 0, 0.9],
        size: Number.isFinite(baseSize) ? Math.max(baseSize, 10) : 10,
        outline: { color: [255, 255, 255, 0.95], width: 1.6 }
      };
    }

    if (graphic.geometry.type === 'polygon') {
      return {
        type: 'simple-fill',
        color: [255, 140, 0, 0.22],
        outline: { color: [255, 140, 0, 0.98], width: 2 }
      };
    }

    return null;
  }

  function clearHoveredSketchGraphic() {
    if (hoveredSketchGraphic && hoveredSketchOriginalSymbol) {
      hoveredSketchGraphic.symbol = hoveredSketchOriginalSymbol;
    }

    hoveredSketchGraphic = null;
    hoveredSketchOriginalSymbol = null;
  }

  function setHoveredSketchGraphic(graphic) {
    if (hoveredSketchGraphic === graphic) {
      return;
    }

    clearHoveredSketchGraphic();
    if (!graphic) {
      return;
    }

    const hoverSymbol = getHoverSymbolForGraphic(graphic);
    if (!hoverSymbol) {
      return;
    }

    hoveredSketchOriginalSymbol = cloneGraphicSymbol(graphic.symbol);
    hoveredSketchGraphic = graphic;
    hoveredSketchGraphic.symbol = hoverSymbol;
  }

  view.on('pointer-move', function (event) {
    if (isMapInteractionBusy()) {
      hoverHitTestRequestId += 1;
      clearHoveredSketchGraphic();
      return;
    }

    const requestId = ++hoverHitTestRequestId;
    view.hitTest(event, { include: [savedSketchesLayer, sketchGraphicsLayer] })
      .then(function (hitResult) {
        if (requestId !== hoverHitTestRequestId) {
          return;
        }

        const results = hitResult && Array.isArray(hitResult.results) ? hitResult.results : [];
        const hit = results.find(function (result) {
          const layer = result.graphic && result.graphic.layer;
          return layer === savedSketchesLayer || layer === sketchGraphicsLayer;
        });

        if (hit && hit.graphic) {
          setHoveredSketchGraphic(hit.graphic);
          view.cursor = 'pointer';
          return;
        }

        clearHoveredSketchGraphic();
        if (!isMapInteractionBusy()) {
          view.cursor = 'auto';
        }
      })
      .catch(function () {
        if (requestId !== hoverHitTestRequestId) {
          return;
        }

        clearHoveredSketchGraphic();
        if (!isMapInteractionBusy()) {
          view.cursor = 'auto';
        }
      });
  });

  view.on('pointer-leave', function () {
    hoverHitTestRequestId += 1;
    clearHoveredSketchGraphic();
    if (!isMapInteractionBusy()) {
      view.cursor = 'auto';
    }
  });

  view.on('click', function (event) {
    if (isMapInteractionBusy()) {
      return;
    }

    view.hitTest(event, { include: [savedSketchesLayer] })
      .then(function (hitResult) {
        const results = hitResult && Array.isArray(hitResult.results) ? hitResult.results : [];
        const hit = results.find(function (result) {
          return result.graphic && result.graphic.layer === savedSketchesLayer;
        });

        if (!hit || !hit.graphic) {
          return;
        }

        const record = findSavedSketchRecordByGraphic(hit.graphic);
        if (!record) {
          return;
        }

        Promise.resolve(analyzeSavedSketchRecord(record, { zoom: false })).catch(function (err) {
          console.error('Saved-sketch click analysis error:', err);
        });
      })
      .catch(function (err) {
        console.error('Saved-sketch click hitTest error:', err);
      });
  });

  const distanceUnitSelect = document.getElementById('distanceUnit');
  const activateDistanceBtn = document.getElementById('activateDistance');
  const clearDistanceBtn = document.getElementById('clearDistance');
  const distanceResult = document.getElementById('distanceResult');

  const areaUnitSelect = document.getElementById('areaUnit');
  const activateAreaBtn = document.getElementById('activateArea');
  const clearAreaBtn = document.getElementById('clearArea');
  const areaResult = document.getElementById('areaResult');

  // Distance Measurement
  activateDistanceBtn.addEventListener('click', function () {
    distanceModeActive = !distanceModeActive;
    if (distanceModeActive) {
      // Deactivate area mode if active
      if (areaModeActive) {
        areaModeActive = false;
        activateAreaBtn.textContent = 'Measure Area';
        activateAreaBtn.style.background = '#17365d';
        clearAreaBtn.style.display = 'none';
      }
      activateDistanceBtn.textContent = 'Stop Measuring Distance';
      activateDistanceBtn.style.background = '#d7191c';
      clearDistanceBtn.style.display = 'inline-block';
      distanceResult.textContent = 'Click on map to add points';
      view.cursor = 'crosshair';
      distancePoints = [];
      measurementGraphicsLayer.removeAll();
    } else {
      activateDistanceBtn.textContent = 'Measure Distance';
      activateDistanceBtn.style.background = '#17365d';
      clearDistanceBtn.style.display = 'none';
      view.cursor = 'auto';
    }
  });

  clearDistanceBtn.addEventListener('click', function () {
    distancePoints = [];
    measurementGraphicsLayer.removeAll();
    distanceResult.textContent = '';
    distanceResult.innerHTML = '';
  });

  // Area Measurement
  activateAreaBtn.addEventListener('click', function () {
    areaModeActive = !areaModeActive;
    if (areaModeActive) {
      // Deactivate distance mode if active
      if (distanceModeActive) {
        distanceModeActive = false;
        activateDistanceBtn.textContent = 'Measure Distance';
        activateDistanceBtn.style.background = '#17365d';
        clearDistanceBtn.style.display = 'none';
      }
      activateAreaBtn.textContent = 'Stop Measuring Area';
      activateAreaBtn.style.background = '#d7191c';
      clearAreaBtn.style.display = 'inline-block';
      areaResult.textContent = 'Click on map to add points. Double-click to finish.';
      view.cursor = 'crosshair';
      areaPoints = [];
      measurementGraphicsLayer.removeAll();
    } else {
      activateAreaBtn.textContent = 'Measure Area';
      activateAreaBtn.style.background = '#17365d';
      clearAreaBtn.style.display = 'none';
      view.cursor = 'auto';
    }
  });

  clearAreaBtn.addEventListener('click', function () {
    areaPoints = [];
    measurementGraphicsLayer.removeAll();
    areaResult.textContent = '';
    areaResult.innerHTML = '';
  });

  // Map click handler for measurements
  let lastClickTime = 0;

  view.on('click', function (event) {
    const currentTime = new Date().getTime();
    const isDoubleClick = currentTime - lastClickTime < 300;
    lastClickTime = currentTime;

    if (distanceModeActive) {
      console.log('Distance mode active, click detected');
      const point = event.mapPoint;
      console.log('Distance point:', point.longitude, point.latitude);
      distancePoints.push(point);

      // Add point graphic
      const pointGraphic = new Graphic({
        geometry: point,
        symbol: {
          type: 'simple-marker',
          color: '#0096c8',
          size: 8,
          outline: { color: 'white', width: 2 }
        }
      });
      measurementGraphicsLayer.add(pointGraphic);

      if (distancePoints.length > 1) {
        // Remove previous line if exists
        const existingLines = measurementGraphicsLayer.graphics.filter(function (g) {
          return g.geometry.type === 'polyline';
        });
        existingLines.forEach(function (g) {
          measurementGraphicsLayer.remove(g);
        });

        // Draw line
        console.log('Creating line with points:', distancePoints);
        
        // Try creating polyline with view's SR
        const polyline = new Polyline({
          paths: [distancePoints.map(p => [p.x, p.y])],
          spatialReference: view.spatialReference
        });
        
        console.log('Created polyline with paths:', polyline.paths, 'SR:', polyline.spatialReference);

        const lineGraphic = new Graphic({
          geometry: polyline,
          symbol: {
            type: 'simple-line',
            color: '#0000ff', // Changed to bright blue
            width: 10, // Increased to 10px
            style: 'solid'
          }
        });
        
        console.log('Created line graphic:', lineGraphic);
        measurementGraphicsLayer.add(lineGraphic);
        console.log('Added distance line graphic, total graphics in layer:', measurementGraphicsLayer.graphics.length);
        
        // Log all graphics in the layer
        measurementGraphicsLayer.graphics.forEach(function(g, i) {
          console.log('Graphic ' + i + ':', g.geometry.type, g);
        });

        // Calculate distance
        let totalDistance = 0;
        for (let i = 0; i < distancePoints.length - 1; i++) {
          totalDistance += geometryEngine.distance(distancePoints[i], distancePoints[i + 1], 'kilometers');
        }

        let displayDistance = totalDistance;
        let unit = distanceUnitSelect.value;

        if (unit === 'miles') {
          displayDistance = totalDistance * 0.621371;
        } else if (unit === 'meters') {
          displayDistance = totalDistance * 1000;
        }

        const unitLabel = unit === 'kilometers' ? 'km' : (unit === 'miles' ? 'mi' : 'm');
        distanceResult.innerHTML = '<strong>Distance: ' + displayDistance.toFixed(2) + ' ' + unitLabel + '</strong>';
      } else {
        distanceResult.textContent = 'Points: 1 - Click to add more points';
      }
    }

    if (areaModeActive) {
      console.log('Area mode active, click detected');
      const point = event.mapPoint;
      console.log('Area point:', point.longitude, point.latitude);

      if (!isDoubleClick) {
        areaPoints.push(point);

        // Add point graphic
        const pointGraphic = new Graphic({
          geometry: point,
          symbol: {
            type: 'simple-marker',
            color: '#00a65d',
            size: 8,
            outline: { color: 'white', width: 2 }
          }
        });
        measurementGraphicsLayer.add(pointGraphic);

        if (areaPoints.length >= 2) {
          // Draw line connecting all points
          const lineCoords = areaPoints.map(function (p) {
            return [p.x, p.y];
          });
          const polyline = new Polyline({
            paths: [lineCoords],
            spatialReference: view.spatialReference
          });

          // Remove previous line if exists
          const existingLines = measurementGraphicsLayer.graphics.filter(function (g) {
            return g.geometry.type === 'polyline';
          });
          existingLines.forEach(function (g) {
            measurementGraphicsLayer.remove(g);
          });

          const lineGraphic = new Graphic({
            geometry: polyline,
            symbol: {
              type: 'simple-line',
              color: '#00ff00', // Changed to bright green
              width: 5, // Increased to 5px
              style: 'solid'
            }
          });
          measurementGraphicsLayer.add(lineGraphic);
        }

        if (areaPoints.length >= 3) {
          // Draw polygon
          const polygonCoords = areaPoints.map(function (p) {
            return [p.x, p.y];
          });
          polygonCoords.push([areaPoints[0].x, areaPoints[0].y]); // Close polygon

          const polygon = new Polygon({
            rings: [polygonCoords],
            spatialReference: view.spatialReference
          });

          // Remove previous polygon if exists
          const existingPolygons = measurementGraphicsLayer.graphics.filter(function (g) {
            return g.geometry.type === 'polygon';
          });
          existingPolygons.forEach(function (g) {
            measurementGraphicsLayer.remove(g);
          });

          const polygonGraphic = new Graphic({
            geometry: polygon,
            symbol: {
              type: 'simple-fill',
              color: 'rgba(0, 165, 93, 0.2)',
              outline: {
                color: '#00a65d',
                width: 2
              }
            }
          });
          measurementGraphicsLayer.add(polygonGraphic);

          // Calculate area
          const areaGeom = new Polygon({
            rings: [polygonCoords],
            spatialReference: view.spatialReference
          });
          let areaValue = geometryEngine.geodesicArea(areaGeom, 'square-kilometers');
          let displayArea = areaValue;
          let unit = areaUnitSelect.value;

          if (unit === 'squareMiles') {
            displayArea = areaValue * 0.386102;
          } else if (unit === 'squareMeters') {
            displayArea = areaValue * 1000000;
          }

          const unitLabel = unit === 'squareKilometers' ? 'km²' : (unit === 'squareMiles' ? 'mi²' : 'm²');
          areaResult.innerHTML = '<strong>Area: ' + displayArea.toFixed(2) + ' ' + unitLabel + '</strong><br>Points: ' + areaPoints.length;
        } else if (areaPoints.length === 2) {
          areaResult.textContent = 'Points: 2 - Click to add more points to create area';
        } else {
          areaResult.textContent = 'Points: ' + areaPoints.length + ' - Click to add more points (need minimum 3)';
        }
      } else if (isDoubleClick && areaPoints.length > 2) {
        // Finish area measurement on double-click
        areaModeActive = false;
        activateAreaBtn.textContent = 'Measure Area';
        activateAreaBtn.style.background = '#17365d';
        clearAreaBtn.style.display = 'none';
        view.cursor = 'auto';
        areaResult.textContent = areaResult.innerHTML + '<br>(Measurement complete - click Clear Area to start new measurement)';
      }
    }
  });

  function showFeatureDetails(features, layerName) {
    if (features.length === 0) return;
    
    // Set title
    detailsTitle.textContent = layerName + ' - Feature Details (' + features.length + ' features)';
    
    // Clear previous content
    detailsTableHead.innerHTML = '';
    detailsTableBody.innerHTML = '';
    
    // Get all unique attribute names
    const allAttributes = new Set();
    features.forEach(feature => {
      Object.keys(feature.attributes).forEach(attr => {
        if (feature.attributes[attr] !== null && feature.attributes[attr] !== undefined) {
          allAttributes.add(attr);
        }
      });
    });
    
    const attributeNames = Array.from(allAttributes);
    
    // Create table header
    const headerRow = document.createElement('tr');
    attributeNames.forEach(attr => {
      const th = document.createElement('th');
      th.textContent = attr;
      headerRow.appendChild(th);
    });
    detailsTableHead.appendChild(headerRow);
    
    // Create table rows for each feature
    features.forEach((feature, index) => {
      const row = document.createElement('tr');
      attributeNames.forEach(attr => {
        const cell = document.createElement('td');
        const value = feature.attributes[attr];
        cell.textContent = value !== null && value !== undefined ? value : '';
        row.appendChild(cell);
      });
      detailsTableBody.appendChild(row);
    });
    
    // Show details section
    featureDetails.classList.remove('hidden');
    
    // Expand popup if collapsed and adjust height
    if (!popupExpanded) {
      togglePopupBtn.click();
    } else {
      bufferPopup.style.maxHeight = '80vh';
    }
  }

  // Toggle popup expand/collapse
  let popupExpanded = true;
  togglePopupBtn.addEventListener('click', function() {
    popupExpanded = !popupExpanded;
    if (popupExpanded) {
      bufferPopup.style.maxHeight = featureDetails.classList.contains('hidden') ? '60vh' : '80vh';
      togglePopupBtn.textContent = '−';
    } else {
      bufferPopup.style.maxHeight = '60px';
      togglePopupBtn.textContent = '+';
    }
  });

  // Close details
  closeDetailsBtn.addEventListener('click', function() {
    featureDetails.classList.add('hidden');
    if (popupExpanded) {
      bufferPopup.style.maxHeight = '60vh';
    }
  });
});
