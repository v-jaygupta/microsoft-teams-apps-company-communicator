// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import * as React from 'react';
import { useTranslation } from 'react-i18next';
import { Spinner } from '@fluentui/react-components';
import { GetScheduledMessagesAction } from '../../actions';
import { RootState, useAppDispatch, useAppSelector } from '../../store';
import { ScheduledMessageDetail } from './scheduledMessageDetail';

export const ScheduledMessages = () => {
  const { t } = useTranslation();
  const scheduledMessages = useAppSelector((state: RootState) => state.messages).scheduledMessages.payload;
  const loader = useAppSelector((state: RootState) => state.messages).isScheduledMessagesFetchOn.payload;
  const dispatch = useAppDispatch();

  React.useEffect(() => {
    if (scheduledMessages && scheduledMessages.length === 0) {
      GetScheduledMessagesAction(dispatch);
    }
  }, [scheduledMessages]);

  return (
    <>
      {loader && <Spinner labelPosition='below' label='Fetching...' />}
      {scheduledMessages && scheduledMessages.length === 0 && !loader && <div>{t('EmptyScheduledMessages')}</div>}
      {scheduledMessages && scheduledMessages.length > 0 && !loader &&
        <ScheduledMessageDetail scheduledMessages={scheduledMessages} />}
    </>
  );
};
