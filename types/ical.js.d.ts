declare module "ical.js" {
  export class Component {
    getAllSubcomponents(name: string): Component[];
    getFirstPropertyValue(name: string): string | null;
    constructor(jCal: any);
  }

  export class Event {
    summary: string;
    description: string;
    location: string;
    startDate: Time | null;
    endDate: Time | null;
    constructor(component: Component);
  }

  export class Time {
    toJSDate(): Date;
  }

  export function parse(input: string): any;
}

